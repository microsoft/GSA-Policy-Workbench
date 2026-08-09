/**
 * Private Access & Application Proxy definitions — Tier 2a.
 *
 * Single source of truth for the V1+ Private Access policy domain. These objects
 * live on DIFFERENT Graph surfaces than the `networkAccess.*` Internet Access
 * tree:
 *
 *   • Private Access apps  → /v1.0/servicePrincipals (tagged)
 *   • App Proxy apps       → /beta/applications (onPremisesPublishing)
 *   • Auth-strength        → /v1.0/identity/conditionalAccess/authenticationStrength
 *
 * In V0 the SPA does NOT call these endpoints live (they need extra scopes the
 * SPA must not request — spec §6.5). They are surfaced read-only from a fixture
 * exported by testharness/Export-GsaFixture.ps1, which captures them as
 * top-level sibling arrays of the filtering-profiles `value`.
 *
 * API version: v1.0 (servicePrincipals, CA) / beta (applications)
 */

// ---------------------------------------------------------------------------
// Private Access application (tagged service principal)
// ---------------------------------------------------------------------------

/** Tag that marks a per-app Private Access (Global Secure Access) application. */
export const PRIVATE_ACCESS_TAG = 'PrivateAccessNonWebApplication';
/** Tag that marks a Quick Access application. */
export const QUICK_ACCESS_TAG = 'NetworkAccessQuickAccessApplication';

// ---------------------------------------------------------------------------
// Application segment (per-app network destination)
// ---------------------------------------------------------------------------

/**
 * The destination shape of an application segment (ztspecs 25395):
 *   • `ip`           single host address
 *   • `ipRange`      inclusive `start..end` range
 *   • `ipRangeCidr`  CIDR block (e.g. `10.0.0.0/24`)
 *   • `fqdn`         single fully-qualified host name
 *   • `dnsSuffix`    wildcard DNS suffix (Private DNS — ztspecs 25399/25400)
 */
export type SegmentDestinationType =
  | 'ip'
  | 'ipRange'
  | 'ipRangeCidr'
  | 'fqdn'
  | 'dnsSuffix'
  | string;

/**
 * One network application segment, read from
 * `/beta/applications/{id}/onPremisesPublishing/segmentsConfiguration/`
 * `microsoft.graph.ipSegmentConfiguration/applicationSegments`.
 *
 * `ports` uses `start-end` strings (e.g. `"445-445"`, `"49152-65535"`); the
 * legacy scalar `port` is `0` when `ports` is populated. `protocol` is
 * `tcp` | `udp` | `tcp,udp` (or `0` for a DNS-suffix segment with no ports).
 */
export interface ApplicationSegment {
  id?: string;
  destinationHost?: string;
  destinationType?: SegmentDestinationType;
  port?: number;
  ports?: string[];
  protocol?: string;
  action?: string;
  exclusions?: unknown;
  inclusions?: unknown;
}

/**
 * A user/group app-role assignment on a Private/Quick Access service principal
 * (`appRoleAssignedTo`). Used to show who can reach the app — ztspecs
 * 25480/25481.
 */
export interface AppRoleAssignment {
  principalId?: string;
  principalType?: string;
  principalDisplayName?: string;
}

/**
 * A Private Access / Quick Access application, projected from a tagged
 * `servicePrincipal`. Both the application object id and the SP object id are
 * kept: `applicationObjectId` is the key for the segments/publishing endpoints
 * (`/beta/applications/{id}/...`); `id` is set to the SP object id when
 * `Directory.Read.All` is available (for assignments), otherwise falls back to
 * the application object id. `appId` is the client id that Conditional Access
 * targets — keep all three.
 */
export interface PrivateAccessApp {
  '@odata.type'?: string;
  /**
   * SP object id when Directory.Read.All is granted, otherwise the application
   * object id. Used as the React key and for /appRoleAssignedTo calls.
   */
  id: string;
  /** Application (client) id — the value CA `includeApplications` targets. */
  appId?: string;
  displayName?: string;
  /**
   * Tags that classify the app. Presence of `PrivateAccessNonWebApplication`
   * vs `NetworkAccessQuickAccessApplication` distinguishes the two kinds.
   */
  tags?: string[];
  /**
   * Custom security attributes — only present when the exporter had
   * CustomSecAttributeAssignment.Read.All + an attribute-reader role. Optional;
   * absence is normal (degrade, don't fail).
   */
  customSecurityAttributes?: Record<string, unknown> | null;
  /**
   * Application object id (`/beta/applications`). Differs from the SP `id` and
   * is the key used to read application segments. Populated by the exporter
   * (correlated by `appId`); absent when not resolved.
   */
  applicationObjectId?: string;
  /**
   * Network application segments — the per-app destinations (FQDNs, IPs, ranges,
   * DNS suffixes) and their ports/protocols. Read by the exporter from the beta
   * `onPremisesPublishing/segmentsConfiguration` route (ztspecs 25395).
   */
  applicationSegments?: ApplicationSegment[];
  /**
   * Whether Private DNS resolution is enabled on the app
   * (`onPremisesPublishing.isDnsResolutionEnabled` — ztspecs 25399). Chiefly
   * meaningful for the Quick Access application.
   */
  isDnsResolutionEnabled?: boolean;
  /**
   * Whether explicit user/group assignment is required to reach the app
   * (`servicePrincipal.appRoleAssignmentRequired`). When `false`, all users
   * have implicit access (ztspecs 25480/25481).
   */
  appRoleAssignmentRequired?: boolean;
  /** Assigned users/groups (`servicePrincipal.appRoleAssignedTo`). */
  appRoleAssignedTo?: AppRoleAssignment[];
}

/** The kind of a Private Access app, derived from its tags. */
export type PrivateAccessKind = 'privateAccess' | 'quickAccess' | 'unknown';

export function privateAccessKind(app: PrivateAccessApp): PrivateAccessKind {
  const tags = app.tags ?? [];
  if (tags.includes(QUICK_ACCESS_TAG)) return 'quickAccess';
  if (tags.includes(PRIVATE_ACCESS_TAG)) return 'privateAccess';
  return 'unknown';
}

// ---------------------------------------------------------------------------
// Application Proxy application
// ---------------------------------------------------------------------------

/**
 * On-premises publishing settings for an App Proxy application. The
 * `externalAuthenticationType` is the security-relevant field:
 *   • `aadPreAuthentication` — Entra pre-auth enforced
 *   • `passthru`            — anonymous reachable (read accurately only from
 *                             the per-app call; spec §6.2)
 */
export interface OnPremisesPublishing {
  externalAuthenticationType?: 'aadPreAuthentication' | 'passthru' | string;
  internalUrl?: string;
  externalUrl?: string;
  isOnPremPublishingEnabled?: boolean;
  [key: string]: unknown;
}

/** An Application Proxy application (`/beta/applications`). */
export interface AppProxyApp {
  '@odata.type'?: string;
  id: string;
  appId?: string;
  displayName?: string;
  onPremisesPublishing?: OnPremisesPublishing;
}

// ---------------------------------------------------------------------------
// Authentication strength policy
// ---------------------------------------------------------------------------

/**
 * An authentication-strength policy referenced by a CA policy's
 * `grantControls.authenticationStrength.id`. Built-in ids:
 *   ...0002 MFA · ...0003 passwordless MFA · ...0004 phishing-resistant MFA.
 */
export interface AuthenticationStrengthPolicy {
  '@odata.type'?: string;
  id: string;
  displayName?: string;
  description?: string;
  policyType?: 'builtIn' | 'custom' | string;
  requirementsSatisfied?: string;
  allowedCombinations?: string[];
}

// ---------------------------------------------------------------------------
// Definitions
// ---------------------------------------------------------------------------

export const PrivateAccessAppDefinition = {
  odataType: '#microsoft.graph.servicePrincipal',
  displayName: 'Private Access App',

  properties: {
    id:                       { label: 'Object ID' },
    appId:                    { label: 'Application ID' },
    displayName:              { label: 'Application Name' },
    tags:                     { label: 'Tags' },
    customSecurityAttributes: { label: 'Custom Security Attributes' },
  },

  // Read-only V1+. Advanced query: needs `ConsistencyLevel: eventual` + $count.
  operations: {
    list: {
      method: 'GET',
      urlTemplate: '/servicePrincipals',
      query:
        "$filter=(tags/any(t:t eq 'PrivateAccessNonWebApplication')" +
        " or tags/any(t:t eq 'NetworkAccessQuickAccessApplication'))" +
        '&$select=id,displayName,appId,tags,customSecurityAttributes&$count=true',
    },
    // Resolve the matching application object (beta) by appId — its object id is
    // the key for reading segments. ztspecs 25395 Q1 / 25400 Q2.
    listApplications: {
      method: 'GET',
      urlTemplate: '/applications',
      query:
        "$filter=(tags/any(t:t eq 'PrivateAccessNonWebApplication')" +
        " or tags/any(t:t eq 'NetworkAccessQuickAccessApplication'))" +
        '&$select=id,appId,displayName,tags',
    },
    // Per-app network segments (beta). ztspecs 25395 Q2 / 25399 Q3 / 25400 Q3.
    listSegments: {
      method: 'GET',
      urlTemplate:
        '/applications/{id}/onPremisesPublishing/segmentsConfiguration/' +
        'microsoft.graph.ipSegmentConfiguration/applicationSegments',
    },
    // On-prem publishing config — carries `isDnsResolutionEnabled` (ztspecs 25399 Q2).
    getPublishing: {
      method: 'GET',
      urlTemplate: '/applications/{id}/onPremisesPublishing',
    },
    // User/group assignments on the service principal (ztspecs 25480/25481).
    getAssignments: {
      method: 'GET',
      urlTemplate: '/servicePrincipals/{id}',
      query:
        '$select=id,appId,accountEnabled,appRoleAssignmentRequired' +
        '&$expand=appRoleAssignedTo($select=principalId,principalType,principalDisplayName)',
    },
  },
} as const;

export const AppProxyAppDefinition = {
  odataType: '#microsoft.graph.application',
  displayName: 'Application Proxy App',

  properties: {
    id:                   { label: 'Object ID' },
    appId:                { label: 'Application ID' },
    displayName:          { label: 'Application Name' },
    onPremisesPublishing: { label: 'On-Premises Publishing' },
  },

  operations: {
    list: {
      method: 'GET',
      urlTemplate: '/applications',
      query:
        '$select=id,displayName&$filter=onPremisesPublishing/isOnPremPublishingEnabled eq true',
    },
    // Per-app call returns the accurate externalAuthenticationType (spec §6.2).
    get: {
      method: 'GET',
      urlTemplate: '/applications/{id}',
      query: '$select=id,displayName,onPremisesPublishing',
    },
  },
} as const;

export const AuthenticationStrengthPolicyDefinition = {
  odataType: '#microsoft.graph.authenticationStrengthPolicy',
  displayName: 'Authentication Strength',

  properties: {
    id:                   { label: 'ID' },
    displayName:          { label: 'Name' },
    policyType:           { label: 'Type' },
    allowedCombinations:  { label: 'Allowed Methods' },
  },

  operations: {
    get: {
      method: 'GET',
      urlTemplate: '/identity/conditionalAccess/authenticationStrength/policies/{id}',
    },
  },
} as const;
