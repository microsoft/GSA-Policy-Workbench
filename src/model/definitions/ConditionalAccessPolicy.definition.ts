/**
 * Conditional Access Policy definition — Tier 2a.
 *
 * Single source of truth for the Conditional Access Policy shape used by
 * the where-used drawer. Full detail requires the optional Policy.Read.All
 * scope. When the scope is absent, only the LinkedCaPolicy stub
 * (id + displayName) from SecurityProfile.definition.ts is available.
 *
 * Note: This type lives in the microsoft.graph namespace (not
 * microsoft.graph.networkaccess), and uses a different base URL.
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/conditionalaccesspolicy
 * API version: beta
 */

// ---------------------------------------------------------------------------
// Supporting types
// ---------------------------------------------------------------------------

/**
 * A resolved directory object (user or group) referenced by a CA policy's
 * user conditions. Populated by the exporter's `directoryObjects` getByIds
 * lookup so the UI can show display names instead of raw object ids. Absent
 * in live Graph mode (no Directory.Read.All) — the UI then falls back to ids.
 */
export interface DirectoryObjectRef {
  id: string;
  displayName?: string;
  '@odata.type'?: string;
  userPrincipalName?: string;
}

/** User and group targeting conditions. */
export interface CaUserCondition {
  includeUsers?: string[];
  excludeUsers?: string[];
  includeGroups?: string[];
  excludeGroups?: string[];
  includeRoles?: string[];
  excludeRoles?: string[];
}/** Application targeting conditions. */
export interface CaApplicationCondition {
  includeApplications?: string[];
  excludeApplications?: string[];
}

/** Device filter condition. */
export interface CaDeviceFilter {
  mode?: 'include' | 'exclude';
  rule?: string;
}

/** Device conditions. */
export interface CaDeviceCondition {
  deviceFilter?: CaDeviceFilter;
}

/** Platform conditions. */
export interface CaPlatformCondition {
  includePlatforms?: string[];
  excludePlatforms?: string[];
}

/** Full set of conditions that must be met for the policy to apply. */
export interface ConditionalAccessConditionSet {
  users?: CaUserCondition;
  applications?: CaApplicationCondition;
  devices?: CaDeviceCondition;
  platforms?: CaPlatformCondition;
  [key: string]: unknown;
}

/** Grant controls that must be fulfilled to pass the policy. */
export interface ConditionalAccessGrantControls {
  operator?: 'AND' | 'OR';
  builtInControls?: string[];
  [key: string]: unknown;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * GSA-relevant Conditional Access **session controls**.
 *
 * These attach a GSA *security / filtering profile* to traffic matched by the
 * CA policy — i.e. Internet Access / SWG-style enforcement. They are **not**
 * the criteria that authorise access to a specific Private Access enterprise
 * app (that is the policy's *target resource* + grant controls).
 */
export interface ConditionalAccessSessionControls {
  /**
   * Documented Graph session control (Internet Access tutorial): links the CA
   * policy to a GSA security/filtering profile by `profileId`.
   */
  globalSecureAccessFilteringProfile?: {
    isEnabled?: boolean;
    profileId?: string;
  } | null;
  /**
   * Legacy / undocumented sibling of `globalSecureAccessFilteringProfile` seen
   * in some CA JSON. Treated the same way (security-profile attachment) but is
   * **not** a documented Graph contract — do not build guidance on it.
   */
  networkAccessSecurity?: {
    isEnabled?: boolean;
    policyId?: string;
  } | null;
  [key: string]: unknown;
}

/**
 * A Conditional Access policy — fetched for the where-used drawer.
 * Requires Policy.Read.All scope. Without the scope, only the stub
 * (id + displayName) from the profile's conditionalAccessPolicies is shown.
 *
 * Note: uses `displayName` (not `name`) per the Graph schema.
 * Note: uses `modifiedDateTime` (not `lastModifiedDateTime`).
 */
export interface ConditionalAccessPolicy {
  '@odata.type'?: string;
  id: string;
  displayName?: string;
  description?: string;
  state?: 'enabled' | 'disabled' | 'enabledForReportingButNotEnforced';
  createdDateTime?: string;
  modifiedDateTime?: string;
  conditions?: ConditionalAccessConditionSet;
  grantControls?: ConditionalAccessGrantControls;
  /**
   * GSA-relevant session controls. When `globalSecureAccessFilteringProfile`
   * (or the legacy `networkAccessSecurity`) is enabled, the policy attaches a
   * GSA security profile to internet traffic — it is an Internet Access policy,
   * not a Private Access app authorization policy.
   */
  sessionControls?: ConditionalAccessSessionControls | null;
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const ConditionalAccessPolicyDefinition = {
  odataType: '#microsoft.graph.conditionalAccessPolicy',
  displayName: 'Conditional Access Policy',

  properties: {
    id:               { label: 'ID' },
    displayName:      { label: 'Policy Name' },
    description:      { label: 'Description' },
    state:            { label: 'State' },
    createdDateTime:  { label: 'Created' },
    modifiedDateTime: { label: 'Last Modified' },
    conditions:       { label: 'Conditions' },
    grantControls:    { label: 'Grant Controls' },
  },

  /**
   * Note: The CA Policy endpoint is under /identity/, not /networkAccess/.
   * The ConditionalAccess repository requires the Policy.Read.All scope.
   */
  operations: {
    list: { method: 'GET', urlTemplate: '/identity/conditionalAccess/policies' },
    get:  { method: 'GET', urlTemplate: '/identity/conditionalAccess/policies/{id}' },
  },
} as const;
