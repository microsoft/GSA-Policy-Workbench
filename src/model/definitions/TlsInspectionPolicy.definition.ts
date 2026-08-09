/**
 * TLS Inspection Policy definition — Tier 2a.
 *
 * Single source of truth for the TLS Inspection Policy GSA object type.
 * Configures TLS termination (break-and-inspect) for network traffic.
 * A profile can have at most one tlsInspectionPolicyLink.
 *
 * Unlike a web content filtering policy:
 *   • the policy-level default lives in `settings.defaultAction`
 *     ('inspect' | 'bypass'), not a top-level `action`;
 *   • each rule (`tlsInspectionRule`) carries its OWN `action`;
 *   • rule destinations live under `matchingConditions.destinations`, where a
 *     web-category destination uses a `values[]` array of category names.
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-tlsinspectionpolicy
 * API version: beta
 */

// ---------------------------------------------------------------------------
// Supporting types
// ---------------------------------------------------------------------------

/** Policy-level settings controlling default TLS inspection behaviour. */
export interface TlsInspectionPolicySettings {
  /** What happens to traffic not matched by any rule. */
  defaultAction?: 'inspect' | 'bypass';
  [key: string]: unknown;
}

/** A destination inside a TLS inspection rule's matching conditions. */
export interface TlsInspectionRuleDestination {
  '@odata.type'?: string;
  /** Web-category destinations carry category names here. */
  values?: string[];
  /** FQDN / IP style destinations carry a single value here. */
  value?: string;
  [key: string]: unknown;
}

/** The conditions that select traffic for a TLS inspection rule. */
export interface TlsInspectionRuleMatchingConditions {
  destinations?: TlsInspectionRuleDestination[];
  [key: string]: unknown;
}

/**
 * A single TLS inspection rule. Note `action` is PER RULE (the WCF model
 * applies its action at the policy level instead).
 */
export interface TlsInspectionRule {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  priority?: number;
  action?: 'inspect' | 'bypass';
  settings?: { status?: 'enabled' | 'disabled'; [key: string]: unknown };
  matchingConditions?: TlsInspectionRuleMatchingConditions | null;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A TLS Inspection Policy — configures break-and-inspect for HTTPS traffic.
 * Rules (tlsInspectionRule) are in `policyRules`, fetched from
 * /networkAccess/tlsInspectionPolicies/{id}/policyRules and inlined by the
 * exporter (the profile $expand does not return them).
 */
export interface TlsInspectionPolicy {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  /** Read-only. Set by the API. */
  version?: string;
  lastModifiedDateTime?: string;
  settings?: TlsInspectionPolicySettings;
  /** Populated by the exporter from the policy's /policyRules sub-route. */
  policyRules?: TlsInspectionRule[];
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const TlsInspectionPolicyDefinition = {
  odataType: '#microsoft.graph.networkaccess.tlsInspectionPolicy',
  displayName: 'TLS Inspection Policy',

  properties: {
    id:                   { label: 'ID' },
    name:                 { label: 'Policy Name' },
    description:          { label: 'Description' },
    version:              { label: 'Version' },
    lastModifiedDateTime: { label: 'Last Modified' },
    settings:             { label: 'Settings' },
    policyRules:          { label: 'Rules' },
  },

  operations: {
    list:   { method: 'GET',    urlTemplate: '/networkAccess/tlsInspectionPolicies' },
    get:    { method: 'GET',    urlTemplate: '/networkAccess/tlsInspectionPolicies/{id}' },
    create: { method: 'POST',   urlTemplate: '/networkAccess/tlsInspectionPolicies' },
    update: { method: 'PATCH',  urlTemplate: '/networkAccess/tlsInspectionPolicies/{id}' },
    delete: { method: 'DELETE', urlTemplate: '/networkAccess/tlsInspectionPolicies/{id}' },
  },
} as const;
