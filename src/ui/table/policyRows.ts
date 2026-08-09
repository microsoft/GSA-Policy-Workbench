/**
 * Policy row model — Tier 1 (UI).
 *
 * Flattens the Security Profile tree returned by the loader into a flat list
 * of one-row-per-rule records for the unified policy table.
 *
 * Pure functions only — no React, no Graph, no side effects. The tree shape
 * comes from the definition files; this module only projects it for display.
 */

import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type {
  FilteringPolicyLink,
  WebContentFilteringPolicy,
} from '../../model/definitions/WebContentFilteringPolicy.definition';
import type {
  FilteringRule,
  RuleDestination,
} from '../../model/definitions/FilteringRule.definition';
import type {
  TlsInspectionPolicy,
  TlsInspectionRule,
} from '../../model/definitions/TlsInspectionPolicy.definition';
import type {
  PromptPolicy,
  PromptRule,
} from '../../model/definitions/PromptPolicy.definition';
import type {
  FilePolicy,
  FileRule,
} from '../../model/definitions/FilePolicy.definition';
import type {
  CloudFirewallPolicy,
  CloudFirewallRule,
} from '../../model/definitions/CloudFirewallPolicy.definition';
import type { ConditionalAccessPolicy, DirectoryObjectRef } from '../../model/definitions/ConditionalAccessPolicy.definition';
import { resolvePolicyTypeLabel } from '../../model/registry';
import { actionToString } from '../actionColor';

/** A single flattened row of the unified policy table. One row per rule. */
export interface PolicyRow {
  /** Stable React key — unique across the whole table. */
  key: string;

  // Profile context
  profileId: string;
  profileName: string;
  profilePriority: number | null;

  // Policy context (empty when a profile has no policies)
  policyType: string;
  /** True when the policy's @odata.type is not in the registry. */
  policyTypeUnknown: boolean;
  policyName: string;
  policyPriority: number | null;
  action: string;
  /** Link state of the policy within this profile ('enabled' | 'disabled' | ''). */
  policyState: string;
  /** Stable key of the owning policy link (`${profileId}/${linkId}`), '' if none. */
  policyKey: string;

  // Rule (empty when a policy has no rules)
  ruleName: string;
  ruleType: string;
  /** Processing order of the rule within its policy (lower = first). */
  rulePriority: number | null;
  /**
   * Per-rule action — only meaningful for TLS inspection rules, whose action
   * (inspect / bypass) lives on the rule rather than the policy. Falls back to
   * the policy-level `action` for web content filtering rows.
   */
  ruleAction?: string;
  /** Human-readable joined destinations for display + search. */
  destinations: string;
  /**
   * Destinations as discrete entries (one per FQDN / URL / web category / IP),
   * for one-per-line rendering in the expanded rule row. Empty falls back to
   * the joined `destinations` string.
   */
  destinationList: string[];

  /** Linked CA policy display names — searchable, not shown as a column. */
  caPolicyNames: string;
}

/** Format one rule destination into a short display string. */
export function formatDestination(dest: RuleDestination): string {
  const type = dest['@odata.type'];
  switch (type) {
    case '#microsoft.graph.networkaccess.webCategory': {
      const d = dest as { displayName?: string; name?: string };
      return d.displayName ?? d.name ?? '(category)';
    }
    case '#microsoft.graph.networkaccess.fqdn':
    case '#microsoft.graph.networkaccess.url':
    case '#microsoft.graph.networkaccess.ipAddress':
    case '#microsoft.graph.networkaccess.ipSubnet': {
      const d = dest as { value?: string };
      return d.value ?? '';
    }
    case '#microsoft.graph.networkaccess.ipRange': {
      const d = dest as { beginAddress?: string; endAddress?: string };
      return `${d.beginAddress ?? ''}–${d.endAddress ?? ''}`;
    }
    default: {
      // Forward-compatible: show any `value`, else the bare type suffix.
      const d = dest as { value?: string };
      if (typeof d.value === 'string') return d.value;
      return type?.split('.').pop() ?? '(destination)';
    }
  }
}

/** Join a rule's destinations into a single display/search string. */
export function formatDestinations(destinations: RuleDestination[] | undefined): string {
  if (!destinations || destinations.length === 0) return '';
  return destinations.map(formatDestination).join(', ');
}

/** A rule's destinations as discrete display entries (one per destination). */
export function destinationEntries(destinations: RuleDestination[] | undefined): string[] {
  if (!destinations) return [];
  return destinations.map(formatDestination).filter((s) => s !== '');
}

function caNames(profile: SecurityProfile): string {
  return (profile.conditionalAccessPolicies ?? [])
    .map((ca) => ca.displayName ?? ca.id)
    .join(', ');
}

function baseRow(profile: SecurityProfile): Omit<PolicyRow, 'key'> {
  return {
    profileId: profile.id,
    profileName: profile.name ?? '(unnamed profile)',
    profilePriority: profile.priority ?? null,
    policyType: '',
    policyTypeUnknown: false,
    policyName: '',
    policyPriority: null,
    action: '',
    policyState: '',
    policyKey: '',
    ruleName: '',
    ruleType: '',
    rulePriority: null,
    destinations: '',
    destinationList: [],
    caPolicyNames: caNames(profile),
  };
}

function withPolicy(
  row: Omit<PolicyRow, 'key'>,
  link: FilteringPolicyLink,
  profileId: string,
): Omit<PolicyRow, 'key'> {
  const policy = link.policy as WebContentFilteringPolicy | undefined;
  // The link type is the reliable discriminant — live Graph omits @odata.type
  // on the nested filtering policy. Action lives on the policy in live Graph;
  // some fixtures put it on the link, so fall back to the link. TLS inspection
  // policies carry their default under settings.defaultAction instead.
  const label = resolvePolicyTypeLabel(link['@odata.type'], policy?.['@odata.type']);
  const tlsDefault = (policy as TlsInspectionPolicy | undefined)?.settings?.defaultAction;
  return {
    ...row,
    policyType: label ?? 'Unsupported type',
    policyTypeUnknown: label === undefined,
    policyName: policy?.name ?? '',
    policyPriority: link.priority ?? null,
    action: actionToString(policy?.action ?? link.action ?? tlsDefault),
    policyState: link.state ?? '',
    policyKey: `${profileId}/${link.id}`,
  };
}

function withRule(
  row: Omit<PolicyRow, 'key'>,
  rule: FilteringRule,
): Omit<PolicyRow, 'key'> {
  return {
    ...row,
    ruleName: rule.name ?? '(unnamed rule)',
    ruleType: rule.ruleType ?? '',
    rulePriority: rule.priority ?? null,
    destinations: formatDestinations(rule.destinations),
    destinationList: destinationEntries(rule.destinations),
  };
}

/** Join a TLS inspection rule's destinations into a display/search string. */
function formatTlsDestinations(rule: TlsInspectionRule): string {
  return tlsDestinationEntries(rule).join(', ');
}

/** A TLS inspection rule's destinations as discrete entries. */
function tlsDestinationEntries(rule: TlsInspectionRule): string[] {
  const dests = rule.matchingConditions?.destinations ?? [];
  const parts: string[] = [];
  for (const d of dests) {
    if (Array.isArray(d.values)) parts.push(...d.values);
    else if (typeof d.value === 'string') parts.push(d.value);
  }
  return parts.filter((s) => s !== '');
}

/** Derive a rule-type token from a TLS rule's first destination. */
function tlsRuleType(rule: TlsInspectionRule): string {
  const odata = rule.matchingConditions?.destinations?.[0]?.['@odata.type'] ?? '';
  if (odata.includes('webCategory')) return 'webCategory';
  if (odata.includes('fqdn')) return 'fqdn';
  if (odata.includes('ipSubnet')) return 'ipSubnet';
  if (odata.includes('ip')) return 'ipAddress';
  return '';
}

function withTlsRule(
  row: Omit<PolicyRow, 'key'>,
  rule: TlsInspectionRule,
): Omit<PolicyRow, 'key'> {
  return {
    ...row,
    ruleName: rule.name ?? '(unnamed rule)',
    ruleType: tlsRuleType(rule),
    rulePriority: rule.priority ?? null,
    ruleAction: actionToString(rule.action),
    destinations: formatTlsDestinations(rule),
    destinationList: tlsDestinationEntries(rule),
  };
}

/** True when a profile policy link is a TLS inspection policy. */
function isTlsLink(link: FilteringPolicyLink): boolean {
  return (
    link['@odata.type'] === '#microsoft.graph.networkaccess.tlsInspectionPolicyLink' ||
    (link.policy as { '@odata.type'?: string } | undefined)?.['@odata.type'] ===
      '#microsoft.graph.networkaccess.tlsInspectionPolicy'
  );
}

/** Join a prompt rule's conversation schemes (+ scan trigger) for display. */
function formatPromptDestinations(rule: PromptRule): string {
  const schemes = rule.matchingConditions?.conversationSchemes ?? [];
  const parts = schemes
    .map((s) => s.schemeName ?? s.url ?? '')
    .filter((p) => p !== '');
  const base = parts.join(', ');
  const scan = rule.matchingConditions?.scanResult;
  if (scan && base) return `${base} · scan: ${scan}`;
  if (scan) return `scan: ${scan}`;
  return base;
}

/** A prompt rule's schemes (+ scan trigger) as discrete entries. */
function promptDestinationEntries(rule: PromptRule): string[] {
  const schemes = rule.matchingConditions?.conversationSchemes ?? [];
  const parts = schemes
    .map((s) => s.schemeName ?? s.url ?? '')
    .filter((p) => p !== '');
  const scan = rule.matchingConditions?.scanResult;
  if (scan) parts.push(`scan: ${scan}`);
  return parts;
}

function withPromptRule(
  row: Omit<PolicyRow, 'key'>,
  rule: PromptRule,
): Omit<PolicyRow, 'key'> {
  return {
    ...row,
    ruleName: rule.name ?? '(unnamed rule)',
    ruleType: 'prompt',
    rulePriority: rule.priority ?? null,
    ruleAction: actionToString(rule.action),
    destinations: formatPromptDestinations(rule),
    destinationList: promptDestinationEntries(rule),
  };
}

/** True when a profile policy link is a prompt protection policy. */
function isPromptLink(link: FilteringPolicyLink): boolean {
  return (
    link['@odata.type'] === '#microsoft.graph.networkaccess.promptPolicyLink' ||
    (link.policy as { '@odata.type'?: string } | undefined)?.['@odata.type'] ===
      '#microsoft.graph.networkaccess.promptPolicy'
  );
}

/** Cap a long value list for display: first `limit`, then "+K more". */
function capValues(values: string[], limit: number): string {
  if (values.length <= limit) return values.join(', ');
  return `${values.slice(0, limit).join(', ')} +${values.length - limit} more`;
}

/**
 * Format a content (file) rule for the Destinations column: the matched
 * destinations (FQDNs / web categories) followed by the distinguishing
 * content-policy signal — the activity (upload/download) and file content types.
 */
function formatFileDestinations(rule: FileRule): string {
  const mc = rule.matchingConditions ?? undefined;
  const destParts: string[] = [];
  for (const d of mc?.destinations ?? []) {
    const vals = d.values ?? (typeof d.value === 'string' ? [d.value] : []);
    if (vals.length === 0) continue;
    const isCategory = (d['@odata.type'] ?? '').toLowerCase().includes('webcategory');
    if (isCategory && vals.length > 8) destParts.push(`${vals.length} web categories`);
    else destParts.push(capValues(vals, 8));
  }
  const where = destParts.join(' · ');

  const fa = mc?.fileAttributes;
  const attrBits: string[] = [];
  if (fa?.activities) attrBits.push(fa.activities);
  if (fa?.contentTypes && fa.contentTypes.length > 0) {
    attrBits.push(capValues(fa.contentTypes, 4));
  }
  const attrs = attrBits.join(' · ');

  if (where && attrs) return `${where} — ${attrs}`;
  return where || attrs;
}

/**
 * A content (file) rule's destinations as discrete entries — every matched
 * FQDN / web category on its own line, with the activity/content-type context
 * appended as a final line.
 */
function fileDestinationEntries(rule: FileRule): string[] {
  const mc = rule.matchingConditions ?? undefined;
  const entries: string[] = [];
  for (const d of mc?.destinations ?? []) {
    const vals = d.values ?? (typeof d.value === 'string' ? [d.value] : []);
    for (const v of vals) if (v) entries.push(v);
  }
  const fa = mc?.fileAttributes;
  const attrBits: string[] = [];
  if (fa?.activities) attrBits.push(fa.activities);
  if (fa?.contentTypes && fa.contentTypes.length > 0) {
    attrBits.push(capValues(fa.contentTypes, 4));
  }
  if (attrBits.length > 0) entries.push(`— ${attrBits.join(' · ')}`);
  return entries;
}

/** Derive a rule-type token from a content rule's first destination. */
function fileRuleType(rule: FileRule): string {
  const odata = (
    rule.matchingConditions?.destinations?.[0]?.['@odata.type'] ?? ''
  ).toLowerCase();
  if (odata.includes('webcategory')) return 'webCategory';
  if (odata.includes('fqdn')) return 'fqdn';
  if (odata.includes('url')) return 'url';
  if (odata.includes('ip')) return 'ipAddress';
  return '';
}

function withFileRule(
  row: Omit<PolicyRow, 'key'>,
  rule: FileRule,
): Omit<PolicyRow, 'key'> {
  return {
    ...row,
    ruleName: rule.name ?? '(unnamed rule)',
    ruleType: fileRuleType(rule),
    rulePriority: rule.priority ?? null,
    ruleAction: actionToString(rule.action),
    destinations: formatFileDestinations(rule),
    destinationList: fileDestinationEntries(rule),
  };
}

/** True when a profile policy link is a content (file) policy. */
function isFileLink(link: FilteringPolicyLink): boolean {
  return (
    link['@odata.type'] === '#microsoft.graph.networkaccess.filePolicyLink' ||
    (link.policy as { '@odata.type'?: string } | undefined)?.['@odata.type'] ===
      '#microsoft.graph.networkaccess.filePolicy'
  );
}

/** Flatten the destination address values of a cloud firewall rule. */
function cloudFirewallDestinationAddresses(rule: CloudFirewallRule): string[] {
  const addrs = rule.matchingConditions?.destinations?.addresses ?? [];
  const out: string[] = [];
  for (const a of addrs) {
    for (const v of a.values ?? []) if (v) out.push(v);
  }
  return out;
}

/** A short "ports + protocol" context string for a cloud firewall rule. */
function cloudFirewallPortContext(rule: CloudFirewallRule): string {
  const dest = rule.matchingConditions?.destinations;
  const bits: string[] = [];
  const ports = dest?.ports ?? [];
  if (ports.length > 0) bits.push(`:${ports.join(', ')}`);
  if (dest?.protocols) bits.push(dest.protocols);
  return bits.join(' ');
}

/**
 * Format a cloud firewall rule for the Destinations column: the destination
 * addresses (IP / FQDN) followed by the matched ports and transport protocol.
 */
function formatCloudFirewallDestinations(rule: CloudFirewallRule): string {
  const where = capValues(cloudFirewallDestinationAddresses(rule), 8);
  const ctx = cloudFirewallPortContext(rule);
  if (where && ctx) return `${where} — ${ctx}`;
  return where || ctx;
}

/**
 * A cloud firewall rule's destinations as discrete entries — each destination
 * address on its own line, with the port/protocol context as a final line.
 */
function cloudFirewallDestinationEntries(rule: CloudFirewallRule): string[] {
  const entries = [...cloudFirewallDestinationAddresses(rule)];
  const ctx = cloudFirewallPortContext(rule);
  if (ctx) entries.push(`— ${ctx}`);
  return entries;
}

/** Derive a rule-type token from a cloud firewall rule's destination addresses. */
function cloudFirewallRuleType(rule: CloudFirewallRule): string {
  const addrs = rule.matchingConditions?.destinations?.addresses ?? [];
  const hasFqdn = addrs.some((a) =>
    (a['@odata.type'] ?? '').toLowerCase().includes('fqdn'),
  );
  return hasFqdn ? 'fqdn' : 'ipAddress';
}

function withCloudFirewallRule(
  row: Omit<PolicyRow, 'key'>,
  rule: CloudFirewallRule,
): Omit<PolicyRow, 'key'> {
  return {
    ...row,
    ruleName: rule.name ?? '(unnamed rule)',
    ruleType: cloudFirewallRuleType(rule),
    rulePriority: rule.priority ?? null,
    ruleAction: actionToString(rule.action),
    destinations: formatCloudFirewallDestinations(rule),
    destinationList: cloudFirewallDestinationEntries(rule),
  };
}

/** True when a profile policy link is a cloud firewall policy. */
function isCloudFirewallLink(link: FilteringPolicyLink): boolean {
  return (
    link['@odata.type'] === '#microsoft.graph.networkaccess.cloudFirewallPolicyLink' ||
    (link.policy as { '@odata.type'?: string } | undefined)?.['@odata.type'] ===
      '#microsoft.graph.networkaccess.cloudFirewallPolicy'
  );
}

/** Sort a policy's rules by processing priority (lower first; missing last). */
function sortByPriority<T extends { priority?: number }>(rules: T[]): T[] {
  return [...rules].sort(
    (a, b) =>
      (a.priority ?? Number.POSITIVE_INFINITY) -
      (b.priority ?? Number.POSITIVE_INFINITY),
  );
}

/**
 * Flatten the profile tree into table rows (one per rule).
 *
 * A profile with no policies still yields one row, and a
 * policy with no rules still yields one row, so every profile stays visible.
 */
export function buildPolicyRows(profiles: SecurityProfile[]): PolicyRow[] {
  const rows: PolicyRow[] = [];

  for (const profile of profiles) {
    const profileRow = baseRow(profile);
    const links = profile.policies ?? [];

    if (links.length === 0) {
      rows.push({ ...profileRow, key: profile.id });
      continue;
    }

    for (const link of links) {
      const policyRow = withPolicy(profileRow, link, profile.id);

      if (isTlsLink(link)) {
        const tlsPolicy = link.policy as TlsInspectionPolicy | undefined;
        const tlsRules = sortByPriority(tlsPolicy?.policyRules ?? []);
        if (tlsRules.length === 0) {
          rows.push({ ...policyRow, key: `${profile.id}/${link.id}` });
          continue;
        }
        for (const rule of tlsRules) {
          rows.push({
            ...withTlsRule(policyRow, rule),
            key: `${profile.id}/${link.id}/${rule.id}`,
          });
        }
        continue;
      }

      if (isPromptLink(link)) {
        const promptPolicy = link.policy as PromptPolicy | undefined;
        const promptRules = sortByPriority(promptPolicy?.policyRules ?? []);
        if (promptRules.length === 0) {
          rows.push({ ...policyRow, key: `${profile.id}/${link.id}` });
          continue;
        }
        for (const rule of promptRules) {
          rows.push({
            ...withPromptRule(policyRow, rule),
            key: `${profile.id}/${link.id}/${rule.id}`,
          });
        }
        continue;
      }

      if (isFileLink(link)) {
        const filePolicy = link.policy as FilePolicy | undefined;
        const fileRules = sortByPriority(filePolicy?.policyRules ?? []);
        if (fileRules.length === 0) {
          rows.push({ ...policyRow, key: `${profile.id}/${link.id}` });
          continue;
        }
        for (const rule of fileRules) {
          rows.push({
            ...withFileRule(policyRow, rule),
            key: `${profile.id}/${link.id}/${rule.id}`,
          });
        }
        continue;
      }

      if (isCloudFirewallLink(link)) {
        const cfwPolicy = link.policy as CloudFirewallPolicy | undefined;
        const cfwRules = sortByPriority(cfwPolicy?.policyRules ?? []);
        if (cfwRules.length === 0) {
          rows.push({ ...policyRow, key: `${profile.id}/${link.id}` });
          continue;
        }
        for (const rule of cfwRules) {
          rows.push({
            ...withCloudFirewallRule(policyRow, rule),
            key: `${profile.id}/${link.id}/${rule.id}`,
          });
        }
        continue;
      }

      const policy = link.policy as WebContentFilteringPolicy | undefined;
      const rules = sortByPriority(policy?.policyRules ?? []);

      if (rules.length === 0) {
        rows.push({ ...policyRow, key: `${profile.id}/${link.id}` });
        continue;
      }

      for (const rule of rules) {
        rows.push({
          ...withRule(policyRow, rule),
          key: `${profile.id}/${link.id}/${rule.id}`,
        });
      }
    }
  }

  return rows;
}

// ---------------------------------------------------------------------------
// Profile grouping + Conditional Access targeting (collapsible table)
// ---------------------------------------------------------------------------

/** The kind of a resolved CA principal — drives the icon shown in the UI. */
export type PrincipalKind = 'user' | 'group' | 'role' | 'all' | 'guests' | 'none' | 'other';

/** A single resolved principal (user/group/role/token) in CA targeting. */
export interface PrincipalRef {
  /** Raw object id or special token (e.g. "All"). */
  id: string;
  /** Resolved display name, friendly token label, or the raw id as fallback. */
  label: string;
  kind: PrincipalKind;
  mode: 'include' | 'exclude';
}

/** One bucket of a CA policy's user/group/role targeting (resolved). */
export interface CaTargetSet {
  label: string;
  mode: 'include' | 'exclude';
  principals: PrincipalRef[];
}

/** A linked CA policy as shown in the profile header. */
export interface LinkedCaInfo {
  id: string;
  name: string;
  state?: string;
  /**
   * True when full CA detail (conditions) was available — i.e. a fixture with
   * the CA sibling array, or a tenant where Policy.Read.All was granted.
   * False when only the stub (id + displayName) is known.
   */
  hasDetail: boolean;
  /** Non-empty target buckets (include/exclude users, groups, roles). */
  targets: CaTargetSet[];
}

/** Aggregated, de-duplicated principals for a profile (users + groups only). */
export interface ProfilePrincipals {
  include: PrincipalRef[];
  exclude: PrincipalRef[];
}

/** At-a-glance roll-up of a profile's effective rules (for the collapsed header). */
export interface ProfileSummary {
  /** Number of rules under block-action policies. */
  blockRuleCount: number;
  /** Number of rules under allow-action policies. */
  allowRuleCount: number;
  /**
   * Action of the catch-all policy (one whose rule targets fqdn `*`), if any.
   * Null when the profile has no catch-all rule.
   */
  catchAllAction: 'allow' | 'block' | null;
}

/** A Security Profile and the flattened policy/rule rows beneath it. */
export interface ProfileGroup {
  profileId: string;
  profileName: string;
  profilePriority: number | null;
  state: string;
  /** Number of filtering-policy links in the profile. */
  policyCount: number;
  /** Number of rule rows (counts placeholder rows for empty policies as 0). */
  ruleCount: number;
  /** Whether priority marks this as the all-users Baseline profile. */
  isBaseline: boolean;
  caPolicies: LinkedCaInfo[];
  /** Aggregated user/group targeting across all linked CA policies. */
  principals: ProfilePrincipals;
  /** True when at least one linked CA policy carried full detail. */
  hasCaDetail: boolean;
  /** At-a-glance roll-up of the profile's effective rules. */
  summary: ProfileSummary;
  rows: PolicyRow[];
  /** Flattened text of CA names + targeting for global search. */
  caSearchText: string;
}

/** Priority of the GSA Baseline profile (applies to everyone, no CA needed). */
const BASELINE_PRIORITY = 65000;

/** Friendly labels for the special non-GUID tokens Graph uses in conditions. */
const TOKEN_LABEL: Record<string, { label: string; kind: PrincipalKind }> = {
  All: { label: 'All users', kind: 'all' },
  None: { label: 'None', kind: 'none' },
  GuestsOrExternalUsers: { label: 'Guests / external users', kind: 'guests' },
};

type DirectoryLookup = Map<string, DirectoryObjectRef>;

/** Resolve one raw id/token into a display-ready principal. */
function resolvePrincipal(
  rawId: string,
  baseKind: PrincipalKind,
  mode: 'include' | 'exclude',
  directory: DirectoryLookup,
): PrincipalRef {
  const token = TOKEN_LABEL[rawId];
  if (token) return { id: rawId, label: token.label, kind: token.kind, mode };

  const ref = directory.get(rawId);
  const odata = ref?.['@odata.type'] ?? '';
  const kind: PrincipalKind = odata.includes('group')
    ? 'group'
    : odata.includes('user')
      ? 'user'
      : baseKind;

  return {
    id: rawId,
    label: ref?.displayName ?? ref?.userPrincipalName ?? rawId,
    kind,
    mode,
  };
}

function bucket(
  label: string,
  mode: 'include' | 'exclude',
  ids: string[] | undefined,
  baseKind: PrincipalKind,
  directory: DirectoryLookup,
): CaTargetSet[] {
  if (!ids || ids.length === 0) return [];
  return [
    {
      label,
      mode,
      principals: ids.map((id) => resolvePrincipal(id, baseKind, mode, directory)),
    },
  ];
}

/** Project a full CA policy's user/group/role conditions into display buckets. */
function targetsFromDetail(
  ca: ConditionalAccessPolicy,
  directory: DirectoryLookup,
): CaTargetSet[] {
  const users = ca.conditions?.users;
  if (!users) return [];
  return [
    ...bucket('Include users', 'include', users.includeUsers, 'user', directory),
    ...bucket('Exclude users', 'exclude', users.excludeUsers, 'user', directory),
    ...bucket('Include groups', 'include', users.includeGroups, 'group', directory),
    ...bucket('Exclude groups', 'exclude', users.excludeGroups, 'group', directory),
    ...bucket('Include roles', 'include', users.includeRoles, 'role', directory),
    ...bucket('Exclude roles', 'exclude', users.excludeRoles, 'role', directory),
  ];
}

function resolveCaPolicies(
  profile: SecurityProfile,
  detailById: Map<string, ConditionalAccessPolicy>,
  directory: DirectoryLookup,
): LinkedCaInfo[] {
  return (profile.conditionalAccessPolicies ?? []).map((stub) => {
    const detail = detailById.get(stub.id);
    return {
      id: stub.id,
      name: detail?.displayName ?? stub.displayName ?? stub.id,
      state: detail?.state,
      hasDetail: detail !== undefined,
      targets: detail ? targetsFromDetail(detail, directory) : [],
    };
  });
}

/**
 * Aggregate user + group principals (not roles) across a profile's linked CA
 * policies into de-duplicated include / exclude lists for the header.
 */
function aggregatePrincipals(caPolicies: LinkedCaInfo[]): ProfilePrincipals {
  const include = new Map<string, PrincipalRef>();
  const exclude = new Map<string, PrincipalRef>();

  for (const ca of caPolicies) {
    for (const set of ca.targets) {
      for (const p of set.principals) {
        if (p.kind === 'role') continue;
        (p.mode === 'include' ? include : exclude).set(p.id, p);
      }
    }
  }

  return { include: [...include.values()], exclude: [...exclude.values()] };
}

/** Build the effective-rules roll-up shown on a collapsed profile header. */
function summarise(profile: SecurityProfile, rows: PolicyRow[]): ProfileSummary {
  let blockRuleCount = 0;
  let allowRuleCount = 0;
  for (const r of rows) {
    if (r.ruleName === '') continue;
    if (r.action === 'block') blockRuleCount += 1;
    else if (r.action === 'allow') allowRuleCount += 1;
  }

  let catchAllAction: 'allow' | 'block' | null = null;
  for (const link of profile.policies ?? []) {
    const policy = link.policy as WebContentFilteringPolicy | undefined;
    const action = policy?.action ?? link.action;
    const hasCatchAll = (policy?.policyRules ?? []).some((rule) =>
      (rule.destinations ?? []).some(
        (d) =>
          d['@odata.type'] === '#microsoft.graph.networkaccess.fqdn' &&
          (d as { value?: string }).value === '*',
      ),
    );
    if (hasCatchAll && (action === 'allow' || action === 'block')) {
      catchAllAction = action;
    }
  }

  return { blockRuleCount, allowRuleCount, catchAllAction };
}

/**
 * Group the profile tree into collapsible profile buckets, resolving each
 * profile's linked CA policies (and their user/group targeting when detail is
 * available) for the table header.
 */
export function buildProfileGroups(
  profiles: SecurityProfile[],
  caDetails: ConditionalAccessPolicy[] = [],
  directoryObjects: DirectoryObjectRef[] = [],
): ProfileGroup[] {
  const detailById = new Map(caDetails.map((ca) => [ca.id, ca]));
  const directory: DirectoryLookup = new Map(
    directoryObjects.map((o) => [o.id, o]),
  );

  return profiles.map((profile) => {
    const rows = buildPolicyRows([profile]);
    const caPolicies = resolveCaPolicies(profile, detailById, directory);
    const principals = aggregatePrincipals(caPolicies);
    const links = profile.policies ?? [];

    const caSearchText = caPolicies
      .flatMap((ca) => [
        ca.name,
        ...ca.targets.flatMap((t) => t.principals.flatMap((p) => [p.label, p.id])),
      ])
      .join(' ');

    return {
      profileId: profile.id,
      profileName: profile.name ?? '(unnamed profile)',
      profilePriority: profile.priority ?? null,
      state: profile.state ?? '',
      policyCount: links.length,
      ruleCount: rows.filter((r) => r.ruleName !== '').length,
      isBaseline: (profile.priority ?? -1) >= BASELINE_PRIORITY,
      caPolicies,
      principals,
      hasCaDetail: caPolicies.some((ca) => ca.hasDetail),
      summary: summarise(profile, rows),
      rows,
      caSearchText,
    };
  });
}


