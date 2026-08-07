/**
 * Derives structured TraceRecord[] from What-If resolver outputs.
 *
 * Pure projection — no evaluation logic. The resolvers stay pure functions;
 * this module translates their already-computed results into a flat list of
 * human-readable decision steps for the trace drawer.
 */

import type { EffectiveResult, AcquisitionEvaluation } from './effectivePolicy';
import type { PrivateAccessWhatIfResult } from './privateAccessWhatIf';
import { type TraceRecord, type TraceDecision } from '../../app/tracer';

export type { TraceRecord };

/** Optional live-resolved context passed alongside the resolver result. */
export interface WhatIfTraceContext {
  resolvedCategoryName?: string;
  resolvedCategoryGroup?: string;
  userGroups?: { id: string; displayName?: string }[];
}

function step(
  stage: string,
  label: string,
  decision: TraceDecision,
  reason: string,
  detail?: Record<string, unknown>,
): TraceRecord {
  return { channel: 'whatif', stage, label, decision, reason, detail, timestamp: Date.now() };
}

function acquisitionDecision(acq: AcquisitionEvaluation): TraceDecision {
  if (acq.isEscape) return 'escape';
  if (acq.preemptedBy) return 'preempted';
  if (acq.verdict === 'forwarded') return 'pass';
  return 'info';
}

function acquisitionSteps(
  acq: AcquisitionEvaluation,
  trafficLabel: string,
): TraceRecord[] {
  const records: TraceRecord[] = [];

  if (acq.preemptedBy) {
    records.push(step(
      'acquisition',
      `${acq.preemptedBy === 'entra' ? 'Microsoft Entra' : 'Microsoft 365'} pre-emption`,
      'preempted',
      acq.reason,
      { verdict: acq.verdict, preemptedBy: acq.preemptedBy },
    ));
  }

  records.push(step(
    'acquisition',
    `${trafficLabel} acquisition`,
    acquisitionDecision(acq),
    acq.reason,
    {
      verdict: acq.verdict,
      profile: acq.profileName,
      matchedRule: acq.match?.ruleName,
      matchedDestination: acq.match?.destinationLabel,
      serviceGroup: acq.match?.serviceGroup,
      category: acq.match?.category,
    },
  ));

  return records;
}

/** Project resolver results into a flat, ordered list of trace steps. */
export function deriveWhatIfTrace(
  result: EffectiveResult,
  paResult?: PrivateAccessWhatIfResult,
  context?: WhatIfTraceContext,
): TraceRecord[] {
  const records: TraceRecord[] = [];

  if (!result.hasUser && !result.hasDestination) return records;

  // ── Stage 0: Live-resolved context (category + group membership) ─────────
  if (context?.resolvedCategoryName && result.destination) {
    const groupNote = context.resolvedCategoryGroup ? ` (${context.resolvedCategoryGroup})` : '';
    records.push(step(
      'context',
      'Destination category',
      'info',
      `${result.destination} → ${context.resolvedCategoryName}${groupNote}`,
      { category: context.resolvedCategoryName, group: context.resolvedCategoryGroup },
    ));
  }

  if (context?.userGroups && context.userGroups.length > 0) {
    const MAX_NAMED = 8;
    const named = context.userGroups.slice(0, MAX_NAMED).map((g) => g.displayName ?? g.id);
    const overflow = context.userGroups.length - MAX_NAMED;
    const summary = overflow > 0 ? `${named.join(', ')} …and ${overflow} more` : named.join(', ');
    records.push(step(
      'context',
      'Group membership',
      'info',
      `${context.userGroups.length} group${context.userGroups.length === 1 ? '' : 's'}: ${summary}`,
      { groupCount: context.userGroups.length, groups: context.userGroups.map((g) => g.displayName ?? g.id) },
    ));
  }

  // ── Stage 1: Acquisition ─────────────────────────────────────────────────
  if (result.acquisition) {
    records.push(...acquisitionSteps(result.acquisition, 'Internet Access'));
  }
  if (paResult?.acquisition) {
    records.push(...acquisitionSteps(paResult.acquisition, 'Private Access'));
  }

  // ── Stage 2: Profile selection ────────────────────────────────────────────
  for (const profile of result.profiles) {
    if (profile.applicability === 'no') continue;

    const isWinner = result.winner?.profile.profileId === profile.profileId;
    const profileDecision: TraceDecision =
      isWinner ? 'winner' :
      profile.applicability === 'yes' ? 'pass' :
      'info';

    records.push(step(
      'profile',
      profile.profileName,
      profileDecision,
      profile.applicabilityReason,
      {
        priority: profile.priority,
        applicability: profile.applicability,
        isBaseline: profile.isBaseline,
      },
    ));

    // ── Stage 3: Rule evaluation (within applicable profiles) ────────────
    if (result.hasDestination) {
      for (const policy of profile.policies) {
        const isPolicyWinner = isWinner && result.winner?.policy.linkId === policy.linkId;
        const ruleDecision: TraceDecision =
          isPolicyWinner ? 'winner' :
          policy.match ? 'match' :
          'no-match';

        records.push(step(
          'rule',
          `${policy.policyName}`,
          ruleDecision,
          policy.match
            ? `Matched rule "${policy.match.ruleName}" (${policy.match.ruleType}) — ${policy.match.destinationLabel}`
            : 'No rule in this policy covers the destination',
          {
            profile: profile.profileName,
            policyType: policy.policyType,
            action: policy.action || '(none)',
            priority: policy.priority,
          },
        ));

        // Rules after the winner were never evaluated by GSA.
        if (isPolicyWinner) break;
      }
    }
  }

  // ── Private Access segment matches ────────────────────────────────────────
  if (paResult?.hasDestination) {
    if (paResult.matches.length === 0) {
      records.push(step(
        'pa-segment',
        'Private Access',
        'no-match',
        'No Private Access app segment covers this destination',
      ));
    }
    for (const m of paResult.matches) {
      const top = m.hits[0];
      records.push(step(
        'pa-segment',
        m.app.name,
        m.supersededBy ? 'skip' : 'match',
        m.supersededBy
          ? `Superseded by ${m.supersededBy.join(', ')} — Quick Access not applied for this destination`
          : `Matched via ${top.reason}: ${top.segment.destinationHost}`,
        { access: m.access, hits: m.hits.length },
      ));
    }
  }

  return records;
}
