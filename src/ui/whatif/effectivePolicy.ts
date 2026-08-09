/**
 * Effective-policy resolver — Tier 1 (UI), pure logic.
 *
 * Implements the read-only "What-If / effective policy" resolution — it
 * shows rule matches but does not simulate live traffic decisions. It
 * operates entirely on the
 * already-loaded profile tree + Conditional Access detail — it makes NO Graph
 * calls and adds no scopes, so it stays within the V0 read-only boundary.
 *
 * Two independent inputs, either or both:
 *   • user        — which Security Profiles apply to that user (via the linked
 *                   Conditional Access policies' user conditions).
 *   • destination — which filtering rule wins for that FQDN / web category,
 *                   following GSA evaluation order (profile priority, then
 *                   policy priority, first match wins).
 *
 * Honest about limits (no live-traffic simulation, no group-membership data):
 *   • Group-scoped CA targeting resolves to "maybe" — we cannot know membership
 *     without Directory.Read.All graph reads that V0 does not perform.
 *   • Report-only / disabled CA policies are flagged as non-enforcing.
 *
 * Pure functions only — no React, no side effects.
 */

import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type { ConditionalAccessPolicy } from '../../model/definitions/ConditionalAccessPolicy.definition';
import type {
  FilteringPolicyLink,
  WebContentFilteringPolicy,
} from '../../model/definitions/WebContentFilteringPolicy.definition';
import type {
  ForwardingProfile,
  ForwardingRule,
  TrafficForwardingType,
} from '../../model/definitions/ForwardingProfile.definition';
import { matchSegmentDestination, type SegmentRow } from '../private/privateAccessRows';
import type {
  FilteringRule,
  RuleDestination,
} from '../../model/definitions/FilteringRule.definition';
import { resolvePolicyTypeLabel } from '../../model/registry';
import { formatDestination } from '../table/policyRows';
import { actionToString } from '../actionColor';
import entraSystemEndpoints from './entraSystemEndpoints.json';

/** Priority of the GSA Baseline profile (applies to everyone, no CA needed). */
const BASELINE_PRIORITY = 65000;

/** How confidently a profile applies to the chosen user. */
export type Applicability = 'yes' | 'maybe' | 'no' | 'unknown';

export type PolicyAction = 'allow' | 'block' | '';

/** A rule's matching destination against the destination query. */
export interface DestinationMatch {
  ruleId: string;
  ruleName: string;
  ruleType: string;
  /** The specific destination string that matched the query. */
  destinationLabel: string;
}

/** One filtering policy inside a profile, evaluated against the inputs. */
export interface PolicyEvaluation {
  linkId: string;
  policyName: string;
  policyType: string;
  action: PolicyAction;
  priority: number | null;
  /** Set when a destination query was given and a rule in this policy matched. */
  match?: DestinationMatch;
}

/** One Security Profile evaluated against the inputs. */
export interface ProfileEvaluation {
  profileId: string;
  profileName: string;
  priority: number | null;
  isBaseline: boolean;
  /** Whether the profile applies to the chosen user. */
  applicability: Applicability;
  /** Plain-language reason for the applicability verdict. */
  applicabilityReason: string;
  /** Policies in evaluation order (priority ascending). */
  policies: PolicyEvaluation[];
  /** First policy in this profile whose rule matched the destination. */
  firstMatch?: PolicyEvaluation;
}

/** The full resolution result. */
export interface EffectiveResult {
  userId?: string;
  userLabel?: string;
  destination?: string;
  hasUser: boolean;
  hasDestination: boolean;
  /** All profiles in evaluation order; non-applicable ones are flagged. */
  profiles: ProfileEvaluation[];
  /**
   * The overall winning profile + policy for the destination, across the
   * profiles that apply to the user. Only set when a destination was given.
   */
  winner?: { profile: ProfileEvaluation; policy: PolicyEvaluation };
  /**
   * Client-side traffic-forwarding acquisition verdict for the destination
   * (stage 1). Only set when a destination
   * was given. This is evaluated independently of `winner`/`profiles` (which
   * are the cloud-side, stage-2 evaluation) — a destination can be acquired
   * (tunneled) yet still resolve to "no rule match", or vice-versa flagged as
   * an escape even though a cloud-side rule would otherwise have blocked it.
   */
  acquisition?: AcquisitionEvaluation;
}

// ---------------------------------------------------------------------------
// Destination matching
// ---------------------------------------------------------------------------

/** Strip scheme + path from a destination query, leaving the host. */
function toHost(query: string): string {
  const noScheme = query.trim().toLowerCase().replace(/^[a-z][a-z0-9+.-]*:\/\//, '');
  return noScheme.split('/')[0];
}

/** Match a host against an FQDN rule value, honouring `*` and `*.suffix`. */
function fqdnMatches(pattern: string, host: string): boolean {
  const p = pattern.trim().toLowerCase();
  if (p === '*' || p === '') return p === '*';
  if (p.startsWith('*.')) {
    const suffix = p.slice(2);
    return host === suffix || host.endsWith(`.${suffix}`);
  }
  return host === p;
}

/** True when a single rule destination matches the destination query. */
function destinationMatches(dest: RuleDestination, query: string, resolvedCategoryName?: string): boolean {
  const q = query.trim().toLowerCase();
  if (q === '') return false;
  const type = dest['@odata.type'];

  switch (type) {
    case '#microsoft.graph.networkaccess.fqdn': {
      const value = (dest as { value?: string }).value ?? '';
      return fqdnMatches(value, toHost(q));
    }
    case '#microsoft.graph.networkaccess.url': {
      const value = ((dest as { value?: string }).value ?? '').toLowerCase();
      return value !== '' && (q.includes(value) || value.includes(q));
    }
    case '#microsoft.graph.networkaccess.webCategory': {
      const d = dest as { name?: string; displayName?: string };
      const ruleCat = (d.displayName ?? d.name ?? '').toLowerCase();
      if (ruleCat === '') return false;
      // Direct: user typed the category name into the destination field
      if (ruleCat === q || ruleCat.includes(q) || q.includes(ruleCat)) return true;
      // Resolved: live category lookup matched this rule's category (exact match only)
      if (resolvedCategoryName) return ruleCat === resolvedCategoryName.toLowerCase();
      return false;
    }
    default: {
      const value = ((dest as { value?: string }).value ?? '').toLowerCase();
      return value !== '' && value === q;
    }
  }
}

/** Find the first matching destination within a rule, if any. */
function ruleMatch(rule: FilteringRule, query: string, resolvedCategoryName?: string): DestinationMatch | undefined {
  for (const dest of rule.destinations ?? []) {
    if (destinationMatches(dest, query, resolvedCategoryName)) {
      return {
        ruleId: rule.id,
        ruleName: rule.name ?? '(unnamed rule)',
        ruleType: rule.ruleType ?? '',
        destinationLabel: formatDestination(dest),
      };
    }
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Client-side stage: traffic-forwarding acquisition (What-If "stage 1")
//
// Before a destination can ever be evaluated by
// Conditional Access + GSA filtering rules (the rest of this file), it must
// first be *acquired* by a traffic-forwarding profile's acquisition rules —
// otherwise the traffic never reaches GSA at all (a "traffic escape"). This
// is the client-side half of What-If; it reasons over the already-loaded
// `ForwardingProfile[]` (state, priority, linked policy rules) and never
// simulates live client/edge connectivity.
// ---------------------------------------------------------------------------

/** How a destination fared against a traffic-forwarding profile's acquisition rules. */
export type AcquisitionVerdict =
  | 'forwarded'
  | 'bypassed'
  | 'unmatched'
  | 'disabled'
  | 'unknown';

/** The specific acquisition rule that decided the verdict, when one matched. */
export interface AcquisitionMatch {
  ruleId: string;
  ruleName: string;
  action: string;
  destinationLabel: string;
  /** M365 service group name (e.g. "Exchange Online") — set when the rule belongs to a named service group. */
  serviceGroup?: string;
  /** M365 Network Connectivity Principles category (`default` / `optimized` / `allow`). */
  category?: string;
}

/**
 * A higher-precedence workload that claimed a destination before the
 * requested workload's own forwarding profile was even evaluated (strict
 * order: Entra -> M365 -> Private -> Internet). `entra` is a heuristic match against a static, non-Graph
 * endpoint list (`entraSystemEndpoints.json`); `m365` is a real
 * forwarding-profile result. Either way, any cloud-side outcome computed for
 * the *originally requested* workload (Internet or Private Access) is
 * hypothetical -- the traffic never actually reached that workload's
 * evaluation.
 */
export type AcquisitionPreemptingWorkload = 'entra' | 'm365';

export interface AcquisitionEvaluation {
  verdict: AcquisitionVerdict;
  /** True for `bypassed` and `unmatched` — the two "flag this to the admin" states. */
  isEscape: boolean;
  profileName?: string;
  reason: string;
  match?: AcquisitionMatch;
  /** Set when a higher-precedence workload pre-empted the requested one (see `AcquisitionPreemptingWorkload`). */
  preemptedBy?: AcquisitionPreemptingWorkload;
}

/**
 * Adapt a forwarding-rule destination to the shape `matchSegmentDestination`
 * (the Private Access overlap matcher — FQDN wildcard, CIDR, IP range, IP
 * prefix) expects, so acquisition rules reuse that same battle-tested logic
 * instead of a second bespoke matcher. Only the two fields the matcher reads
 * (`destinationHost` + `destinationType`) are meaningful here; the rest are
 * inert filler. `webCategory` destinations aren't supported by the matcher —
 * they're rare on forwarding rules and reported as no-match.
 */
function toAcquisitionSegmentRow(dest: RuleDestination): SegmentRow | null {
  const type = dest['@odata.type'];
  let destinationHost: string;
  let destinationType: string;
  switch (type) {
    case '#microsoft.graph.networkaccess.fqdn':
      destinationType = 'fqdn';
      destinationHost = (dest as { value?: string }).value ?? '';
      break;
    case '#microsoft.graph.networkaccess.url':
      // No dedicated URL destination type in the segment matcher — approximate
      // with the host portion, same as the cloud-side `url` case above.
      destinationType = 'fqdn';
      destinationHost = toHost((dest as { value?: string }).value ?? '');
      break;
    case '#microsoft.graph.networkaccess.ipAddress':
      destinationType = 'ip';
      destinationHost = (dest as { value?: string }).value ?? '';
      break;
    case '#microsoft.graph.networkaccess.ipRange': {
      const r = dest as { beginAddress?: string; endAddress?: string };
      destinationType = 'ipRange';
      destinationHost = `${r.beginAddress ?? ''}..${r.endAddress ?? ''}`;
      break;
    }
    case '#microsoft.graph.networkaccess.ipSubnet':
      destinationType = 'ipRangeCidr';
      destinationHost = (dest as { value?: string }).value ?? '';
      break;
    default:
      return null;
  }
  if (!destinationHost) return null;
  return {
    id: '',
    destinationHost,
    destinationType,
    destinationTypeLabel: '',
    ports: '',
    protocol: '',
    action: '',
    breadth: 'granular',
    isDnsSuffix: false,
    hasPort53: false,
    widePorts: false,
  };
}

/** Find the first acquisition rule whose destinations cover the query. */
function forwardingRuleMatch(
  rule: ForwardingRule,
  query: string,
  serviceGroup?: string,
): AcquisitionMatch | undefined {
  for (const dest of rule.destinations ?? []) {
    const segRow = toAcquisitionSegmentRow(dest);
    if (!segRow) continue;
    if (matchSegmentDestination(segRow, query)) {
      return {
        ruleId: rule.id,
        ruleName: rule.name ?? '(unnamed rule)',
        action: rule.action ?? '',
        destinationLabel: formatDestination(dest),
        serviceGroup,
        category: rule.category,
      };
    }
  }
  return undefined;
}

/**
 * Match a destination against the static Microsoft Entra system-profile
 * endpoint list (`entraSystemEndpoints.json`). The Entra system profile is
 * always processed first, ahead of M365/Private/Internet — but it has no Graph representation at
 * all, so this is a curated heuristic, not tenant data. Returns the matched
 * pattern + note, or `null`.
 */
function matchEntraSystemEndpoint(destination: string): { pattern: string; note?: string } | null {
  const host = toHost(destination);
  if (!host) return null;
  for (const entry of entraSystemEndpoints.endpoints as { pattern: string; note?: string }[]) {
    if (fqdnMatches(entry.pattern, host)) return entry;
  }
  return null;
}

/**
 * Evaluate whether a destination is acquired (tunneled to GSA) by ONE
 * specific forwarding profile (`trafficForwardingType`), ignoring workload
 * precedence — either because the profile is disabled, an acquisition rule
 * explicitly bypasses it, or no acquisition rule covers it at all.
 *
 * `unknown` covers both "no forwarding-profile data in this source" and "the
 * profile has no captured acquisition rules" — distinguishing those from a
 * genuine escape matters: we should never claim traffic escapes GSA just
 * because this tool didn't fetch the rules.
 *
 * Rule precedence: the public `forwardingRule` Graph schema has no numeric
 * `priority` field (confirmed against the beta resource docs, 2026-07-24), so
 * array order cannot be trusted as a tie-break. Instead, ALL `bypass`-action
 * rules across every linked policy are checked before any `forward`-action
 * rule — this reproduces the named "Custom Bypass / Default Bypass beat
 * Default Acquire" precedence without
 * relying on policy names or Graph-returned ordering.
 */
function evaluateAcquisitionForType(
  forwardingProfiles: ForwardingProfile[],
  trafficForwardingType: TrafficForwardingType,
  destination: string,
): AcquisitionEvaluation {
  const profile = forwardingProfiles.find(
    (p) => p.trafficForwardingType === trafficForwardingType,
  );
  if (!profile) {
    return {
      verdict: 'unknown',
      isEscape: false,
      reason: `No ${trafficForwardingType} forwarding-profile data available in this source`,
    };
  }

  if (profile.state === 'disabled') {
    return {
      verdict: 'disabled',
      isEscape: true,
      profileName: profile.name,
      reason: `${profile.name ?? 'The'} forwarding profile is disabled — this traffic class is never sent to Global Secure Access`,
    };
  }

  // Sort links by priority ascending (IA profile only; null = no ordering enforced).
  const links = [...(profile.policies ?? [])].sort((a, b) => {
    if (a.priority == null && b.priority == null) return 0;
    if (a.priority == null) return 1;
    if (b.priority == null) return -1;
    return a.priority - b.priority;
  });

  const bypassRules: { rule: ForwardingRule; serviceGroup?: string }[] = [];
  const forwardRules: { rule: ForwardingRule; serviceGroup?: string }[] = [];
  for (const link of links) {
    const groupName = link.policy?.name;
    for (const rule of link.policy?.policyRules ?? []) {
      // Disabled service group: all its destinations are bypassed regardless of rule action.
      const effectiveBypass = link.state === 'disabled' || rule.action === 'bypass';
      (effectiveBypass ? bypassRules : forwardRules).push({ rule, serviceGroup: groupName });
    }
  }
  if (bypassRules.length === 0 && forwardRules.length === 0) {
    return {
      verdict: 'unknown',
      isEscape: false,
      profileName: profile.name,
      reason: 'Acquisition rules were not captured for this forwarding profile',
    };
  }

  // Bypass rules first, regardless of capture order (see doc comment above).
  for (const { rule, serviceGroup } of [...bypassRules, ...forwardRules]) {
    const match = forwardingRuleMatch(rule, destination, serviceGroup);
    if (!match) continue;
    if (match.action === 'bypass') {
      const groupNote = match.serviceGroup ? ` (${match.serviceGroup})` : '';
      return {
        verdict: 'bypassed',
        isEscape: true,
        profileName: profile.name,
        reason: `Explicitly bypassed by acquisition rule "${match.ruleName}"${groupNote} — not sent to Global Secure Access`,
        match,
      };
    }
    // Disabled group's rules carry action from the original rule but we forced them to bypass above;
    // if we reach here, the link was enabled and action is 'forward'.
    const groupNote = match.serviceGroup ? ` (${match.serviceGroup})` : '';
    return {
      verdict: 'forwarded',
      isEscape: false,
      profileName: profile.name,
      reason: `Acquired by rule "${match.ruleName}"${groupNote}`,
      match,
    };
  }

  return {
    verdict: 'unmatched',
    isEscape: true,
    profileName: profile.name,
    reason: `No acquisition rule in "${profile.name ?? trafficForwardingType}" covers this destination — it may never reach Global Secure Access`,
  };
}

/**
 * Evaluate whether a destination is acquired (tunneled to GSA) or escapes,
 * for the given workload (`internet` or `private`) — honouring GSA's fixed
 * workload precedence **Entra → M365 → Private → Internet**. A destination handled by a
 * higher-precedence workload pre-empts the requested one entirely: the
 * requested workload's own policy evaluation never runs, so any cloud-side
 * outcome computed for it downstream is hypothetical (`preemptedBy` is set).
 *
 * Only `m365` is checked as a pre-emption candidate today (Private Access
 * segment matching pre-empting an Internet Access query is a further
 * workload-order step that is NOT modeled yet (item 5) — because it needs
 * the separately-built Private Access view model,
 * not just `ForwardingProfile` data).
 */
export function evaluateAcquisition(
  forwardingProfiles: ForwardingProfile[],
  trafficForwardingType: TrafficForwardingType,
  destination: string,
): AcquisitionEvaluation {
  const entraMatch = matchEntraSystemEndpoint(destination);
  if (entraMatch) {
    return {
      verdict: 'forwarded',
      isEscape: false,
      profileName: 'Microsoft Entra traffic (system profile)',
      reason: `Matches a known Microsoft Entra identity endpoint (${entraMatch.pattern}) — always tunneled via the system Microsoft Entra profile before ${trafficForwardingType} is evaluated`,
      preemptedBy: 'entra',
    };
  }

  if (trafficForwardingType !== 'm365') {
    const m365 = evaluateAcquisitionForType(forwardingProfiles, 'm365', destination);
    if (m365.verdict === 'forwarded') {
      return {
        ...m365,
        reason: `Handled by the Microsoft 365 traffic profile (${m365.reason}) — evaluated before ${trafficForwardingType === 'internet' ? 'Internet Access' : 'Private Access'}`,
        preemptedBy: 'm365',
      };
    }
    if (m365.verdict === 'bypassed') {
      return {
        ...m365,
        reason: `Microsoft 365 traffic forwarding bypasses this destination (${m365.reason}) — never reaches Global Secure Access via any workload`,
        preemptedBy: 'm365',
      };
    }
  }

  return evaluateAcquisitionForType(forwardingProfiles, trafficForwardingType, destination);
}

// ---------------------------------------------------------------------------
// User applicability (via Conditional Access user conditions)
// ---------------------------------------------------------------------------

function policyAction(link: FilteringPolicyLink): PolicyAction {
  const policy = link.policy as WebContentFilteringPolicy | undefined;
  return actionToString(policy?.action ?? link.action) as PolicyAction;
}

function comparePriority(a: number | null, b: number | null): number {
  if (a === b) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  return a - b;
}

function caApplicability(
  ca: ConditionalAccessPolicy,
  userId: string,
  groupIds?: string[],
): { verdict: Applicability; note: string } {
  if (ca.state === 'disabled') {
    return { verdict: 'no', note: `${name(ca)} is disabled` };
  }
  const reportOnly = ca.state === 'enabledForReportingButNotEnforced';

  const users = ca.conditions?.users;
  if (!users) {
    return { verdict: 'unknown', note: `${name(ca)} has no user conditions` };
  }

  if ((users.excludeUsers ?? []).includes(userId)) {
    return { verdict: 'no', note: `excluded from ${name(ca)}` };
  }

  // Group exclusion: checked before include-all so an excluded group beats a broad include.
  if (groupIds && (users.excludeGroups ?? []).some((g) => groupIds.includes(g))) {
    return { verdict: 'no', note: `excluded from ${name(ca)} via group` };
  }

  const includeUsers = users.includeUsers ?? [];
  if (includeUsers.includes('All') || includeUsers.includes(userId)) {
    return {
      verdict: 'yes',
      note: reportOnly ? `${name(ca)} (report-only)` : name(ca),
    };
  }

  if ((users.includeGroups ?? []).length > 0) {
    if (!groupIds) {
      return { verdict: 'maybe', note: `${name(ca)} targets groups (membership unknown)` };
    }
    if (users.includeGroups!.some((g) => groupIds.includes(g))) {
      return {
        verdict: 'yes',
        note: reportOnly ? `${name(ca)} (report-only, via group)` : `${name(ca)} (via group)`,
      };
    }
    return { verdict: 'no', note: `not in any group targeted by ${name(ca)}` };
  }

  return { verdict: 'no', note: `not targeted by ${name(ca)}` };
}

function name(ca: ConditionalAccessPolicy): string {
  return ca.displayName ?? ca.id;
}

function profileApplicability(
  profile: SecurityProfile,
  detailById: Map<string, ConditionalAccessPolicy>,
  userId: string,
  groupIds?: string[],
): { applicability: Applicability; reason: string } {
  if ((profile.priority ?? -1) >= BASELINE_PRIORITY) {
    return { applicability: 'yes', reason: 'Baseline — applies to all users' };
  }

  const stubs = profile.conditionalAccessPolicies ?? [];
  if (stubs.length === 0) {
    return { applicability: 'no', reason: 'No Conditional Access policy linked' };
  }

  const verdicts = stubs.map((stub) => {
    const detail = detailById.get(stub.id);
    if (!detail) {
      return { verdict: 'unknown' as Applicability, note: `${stub.displayName ?? stub.id} (no detail)` };
    }
    return caApplicability(detail, userId, groupIds);
  });

  const yes = verdicts.find((v) => v.verdict === 'yes');
  if (yes) return { applicability: 'yes', reason: `Included via ${yes.note}` };

  const maybe = verdicts.find((v) => v.verdict === 'maybe');
  if (maybe) return { applicability: 'maybe', reason: maybe.note };

  const unknown = verdicts.find((v) => v.verdict === 'unknown');
  if (unknown) {
    return {
      applicability: 'unknown',
      reason: 'Targeting detail unavailable (Policy.Read.All not granted)',
    };
  }

  return { applicability: 'no', reason: verdicts[0]?.note ?? 'Not targeted' };
}

// ---------------------------------------------------------------------------
// Policy evaluation
// ---------------------------------------------------------------------------

function evaluatePolicies(
  profile: SecurityProfile,
  destination: string | undefined,
  resolvedCategoryName?: string,
): PolicyEvaluation[] {
  const links = [...(profile.policies ?? [])].sort((a, b) =>
    comparePriority(a.priority ?? null, b.priority ?? null),
  );

  return links.map((link) => {
    const policy = link.policy as WebContentFilteringPolicy | undefined;
    const label = resolvePolicyTypeLabel(link['@odata.type'], policy?.['@odata.type']);

    let match: DestinationMatch | undefined;
    if (destination && destination.trim() !== '') {
      // Evaluate rules in processing-priority order (lower first), so the
      // first match reflects the order GSA would apply them.
      const rules = [...(policy?.policyRules ?? [])].sort(
        (a, b) =>
          (a.priority ?? Number.POSITIVE_INFINITY) -
          (b.priority ?? Number.POSITIVE_INFINITY),
      );
      for (const rule of rules) {
        match = ruleMatch(rule, destination, resolvedCategoryName);
        if (match) break;
      }
    }

    return {
      linkId: link.id,
      policyName: policy?.name ?? '(unnamed policy)',
      policyType: label ?? 'Unsupported type',
      action: policyAction(link),
      priority: link.priority ?? null,
      match,
    };
  });
}

// ---------------------------------------------------------------------------
// Public resolver
// ---------------------------------------------------------------------------

export interface ResolveInput {
  userId?: string;
  userLabel?: string;
  destination?: string;
  /** Transitive group ids for the What-If user — resolves group-targeted CA policies from 'maybe' to yes/no. */
  groupIds?: string[];
  /** Web content category resolved for the destination (live mode) — enables category-rule matching by FQDN. */
  resolvedCategoryName?: string;
}

/**
 * Resolve the effective policy for an optional user and/or destination.
 *
 * Profiles are evaluated in GSA order (priority ascending). For a destination,
 * the winner is the first matching policy in the first applicable profile.
 */
export function resolveEffective(
  profiles: SecurityProfile[],
  caDetails: ConditionalAccessPolicy[],
  input: ResolveInput,
  forwardingProfiles: ForwardingProfile[] = [],
): EffectiveResult {
  const detailById = new Map(caDetails.map((ca) => [ca.id, ca]));
  const userId = input.userId?.trim() || undefined;
  const destination = input.destination?.trim() || undefined;
  const hasUser = Boolean(userId);
  const hasDestination = Boolean(destination);

  const ordered = [...profiles].sort((a, b) =>
    comparePriority(a.priority ?? null, b.priority ?? null),
  );

  const evaluations: ProfileEvaluation[] = ordered.map((profile) => {
    const { applicability, reason } = userId
      ? profileApplicability(profile, detailById, userId, input.groupIds)
      : { applicability: 'yes' as Applicability, reason: 'No user filter applied' };

    const policies = evaluatePolicies(profile, destination, input.resolvedCategoryName);
    const firstMatch = policies.find((p) => p.match);

    return {
      profileId: profile.id,
      profileName: profile.name ?? '(unnamed profile)',
      priority: profile.priority ?? null,
      isBaseline: (profile.priority ?? -1) >= BASELINE_PRIORITY,
      applicability,
      applicabilityReason: reason,
      policies,
      firstMatch,
    };
  });

  let winner: EffectiveResult['winner'];
  if (hasDestination) {
    for (const profile of evaluations) {
      // Only enforced ("yes") profiles decide the outcome. "maybe"/"unknown"
      // are surfaced but never claimed as the winner.
      if (profile.applicability !== 'yes') continue;
      if (profile.firstMatch) {
        winner = { profile, policy: profile.firstMatch };
        break;
      }
    }
  }

  return {
    userId,
    userLabel: input.userLabel,
    destination,
    hasUser,
    hasDestination,
    profiles: evaluations,
    winner,
    acquisition: hasDestination
      ? evaluateAcquisition(forwardingProfiles, 'internet', destination!)
      : undefined,
  };
}
