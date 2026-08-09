/**
 * TanStack Query cache keys — centralised to avoid key collisions.
 *
 * Convention: [scope, ...discriminators]
 * All keys are `as const` so TanStack can do exact structural comparison.
 */

export const queryKeys = {
  /** The full Security Profile tree (profiles + linked policies + rules).
   *  Keyed by source id so each data source has its own cache slot —
   *  switching sources never evicts another source's cached data. */
  securityProfiles: (sourceId: string) => ['securityProfiles', sourceId] as const,

  /** All filtering policies in the tenant (flat list, no profile context). */
  filteringPolicies: () => ['filteringPolicies'] as const,

  /** Rules for a specific filtering policy. */
  filteringPolicyRules: (policyId: string) =>
    ['filteringPolicies', policyId, 'rules'] as const,

  /** Full Conditional Access policy detail for a set of ids (where-used drawer). */
  conditionalAccessPolicies: (ids: readonly string[]) =>
    ['conditionalAccessPolicies', ...ids] as const,
} as const;

