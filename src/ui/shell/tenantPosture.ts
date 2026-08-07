/**
 * Tenant posture computation — Tier 1 (UI), pure functions only.
 *
 * Cross-references the tenant-wide policy inventory against the loaded Security
 * Profile tree to derive two read-only inspection signals (spec §10, V2):
 *
 *   - Orphaned policies (item 4): policies that exist in the tenant but are not
 *     linked to any Security Profile.
 *   - Tenant-limit counters (item 6): current usage against the published
 *     Internet Access limits.
 *
 * These are **inspection hints**, never a pass/fail verdict. No React, no
 * Graph, no side effects.
 */

import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type { TenantPolicy } from '../../adapters/graph/loader';
import { policyTypeLabel } from '../../model/registry';
import { POLICY_TYPE_COLOR } from '../policyTypeMeta';
import type { BadgeProps } from '@fluentui/react-components';

/**
 * Published Internet Access limits (per tenant) — see
 * https://learn.microsoft.com/en-us/entra/global-secure-access/reference-current-known-limitations#internet-access-limitations
 */
export const IA_LIMITS = {
  profiles: 256,
  policies: 1000,
  rules: 1000,
  destinations: 8000,
} as const;

/** Fraction of the limit at/above which a counter turns amber. */
const WARN_FRACTION = 0.8;

export type CounterLevel = 'ok' | 'warn' | 'full';

/** One tenant-limit counter ("N of MAX"). */
export interface LimitCounter {
  key: keyof typeof IA_LIMITS;
  label: string;
  count: number;
  max: number;
  level: CounterLevel;
  /** True when the count is derived from loaded (linked) policies only. */
  linkedOnly: boolean;
}

/** A group of orphaned policies sharing one policy type. */
export interface OrphanGroup {
  typeLabel: string;
  color: BadgeProps['color'];
  policies: { id: string; name: string }[];
}

/** The full tenant-posture result surfaced in the strip. */
export interface TenantPosture {
  counters: LimitCounter[];
  orphaned: OrphanGroup[];
  orphanCount: number;
}

/** A minimal projection of a policy's rule for destination counting. */
interface CountableRule {
  destinations?: unknown[];
  matchingConditions?: {
    destinations?: Array<{ values?: unknown[]; value?: unknown }>;
  };
}

/** A policy whose rules we can count (shape varies across policy kinds). */
interface CountablePolicy {
  id?: string;
  policyRules?: CountableRule[];
}

function level(count: number, max: number): CounterLevel {
  if (count >= max) return 'full';
  if (count >= max * WARN_FRACTION) return 'warn';
  return 'ok';
}

/** Count the destinations a single rule targets (across rule kinds). */
function countRuleDestinations(rule: CountableRule): number {
  let n = 0;
  if (Array.isArray(rule.destinations)) n += rule.destinations.length;
  const md = rule.matchingConditions?.destinations;
  if (Array.isArray(md)) {
    for (const d of md) {
      if (Array.isArray(d.values)) n += d.values.length;
      else if (d.value != null) n += 1;
    }
  }
  return n;
}

/**
 * Compute tenant posture (orphaned policies + limit counters).
 *
 * Rule and destination counts are taken from the loaded profile tree, counting
 * each policy once even when it is linked to several profiles. Orphaned
 * policies contribute to the *policy* count (tenant-wide) but, having no rules
 * in the tree, are not reflected in the rule/destination counts — those are
 * flagged `linkedOnly`.
 */
export function computeTenantPosture(
  profiles: SecurityProfile[],
  tenantPolicies: TenantPolicy[],
): TenantPosture {
  // Referenced policy ids + the unique linked policy objects (de-duplicated).
  const referenced = new Set<string>();
  const uniquePolicies = new Map<string, CountablePolicy>();
  for (const profile of profiles) {
    for (const link of profile.policies ?? []) {
      const policy = link.policy as CountablePolicy | undefined;
      const pid = policy?.id;
      if (!pid) continue;
      referenced.add(pid);
      if (!uniquePolicies.has(pid)) uniquePolicies.set(pid, policy as CountablePolicy);
    }
  }

  let rules = 0;
  let destinations = 0;
  for (const policy of uniquePolicies.values()) {
    for (const rule of policy.policyRules ?? []) {
      rules += 1;
      destinations += countRuleDestinations(rule);
    }
  }

  // Tenant-wide policy count: prefer the inventory; fall back to linked count
  // when the inventory is unavailable (e.g. a pre-tenantPolicies fixture).
  const policyCount =
    tenantPolicies.length > 0 ? tenantPolicies.length : uniquePolicies.size;
  const linkedOnly = tenantPolicies.length > 0 && referenced.size < tenantPolicies.length;

  const counters: LimitCounter[] = [
    {
      key: 'profiles',
      label: 'Security profiles',
      count: profiles.length,
      max: IA_LIMITS.profiles,
      level: level(profiles.length, IA_LIMITS.profiles),
      linkedOnly: false,
    },
    {
      key: 'policies',
      label: 'Policies',
      count: policyCount,
      max: IA_LIMITS.policies,
      level: level(policyCount, IA_LIMITS.policies),
      linkedOnly: false,
    },
    {
      key: 'rules',
      label: 'Rules',
      count: rules,
      max: IA_LIMITS.rules,
      level: level(rules, IA_LIMITS.rules),
      linkedOnly,
    },
    {
      key: 'destinations',
      label: 'Destinations',
      count: destinations,
      max: IA_LIMITS.destinations,
      level: level(destinations, IA_LIMITS.destinations),
      linkedOnly,
    },
  ];

  // Orphaned policies — in the tenant inventory but not referenced by any
  // profile. Grouped by policy type for display.
  const groups = new Map<string, OrphanGroup>();
  for (const tp of tenantPolicies) {
    if (referenced.has(tp.id)) continue;
    const typeLabel = policyTypeLabel(tp['@odata.type']) ?? 'Unsupported type';
    let group = groups.get(typeLabel);
    if (!group) {
      group = {
        typeLabel,
        color: POLICY_TYPE_COLOR[typeLabel] ?? 'subtle',
        policies: [],
      };
      groups.set(typeLabel, group);
    }
    group.policies.push({ id: tp.id, name: tp.name || '(unnamed policy)' });
  }

  const orphaned = [...groups.values()].sort((a, b) =>
    a.typeLabel.localeCompare(b.typeLabel),
  );
  const orphanCount = orphaned.reduce((sum, g) => sum + g.policies.length, 0);

  return { counters, orphaned, orphanCount };
}
