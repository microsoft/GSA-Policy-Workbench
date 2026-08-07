/**
 * Threat Intelligence Policy definition — Tier 2a.
 *
 * Single source of truth for the Threat Intelligence Policy GSA object type.
 * Evaluates and enforces security controls based on known threat intelligence
 * (malicious IPs, domains, URLs) via the Global Secure Access service.
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-threatintelligencepolicy
 * API version: beta
 */

// ---------------------------------------------------------------------------
// Supporting types
// ---------------------------------------------------------------------------

/**
 * Policy-level settings for threat intelligence evaluation.
 * Exact shape subject to beta schema change.
 */
export interface ThreatIntelligencePolicySettings {
  [key: string]: unknown;
}

/**
 * A Threat Intelligence policy rule. Typed loosely until a dedicated
 * definition file is added for threatIntelligenceRule.
 */
export interface ThreatIntelligenceRule {
  '@odata.type'?: string;
  id: string;
  name?: string;
  [key: string]: unknown;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A Threat Intelligence Policy — applies security controls based on known
 * threats. Rules are in `policyRules` when expanded.
 *
 * Note: the JSON representation in Graph docs includes a `kind` field
 * (not documented in the properties table) — captured here for completeness.
 */
export interface ThreatIntelligencePolicy {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  /** Read-only. Set by the API. */
  version?: string;
  lastModifiedDateTime?: string;
  /** Undocumented field observed in beta responses — kept for forward compat. */
  kind?: string;
  settings?: ThreatIntelligencePolicySettings;
  /** Populated when expanded via $expand=policyRules. */
  policyRules?: ThreatIntelligenceRule[];
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const ThreatIntelligencePolicyDefinition = {
  odataType: '#microsoft.graph.networkaccess.threatIntelligencePolicy',
  displayName: 'Threat Intelligence Policy',

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
    list:   { method: 'GET',    urlTemplate: '/networkAccess/threatIntelligencePolicies' },
    get:    { method: 'GET',    urlTemplate: '/networkAccess/threatIntelligencePolicies/{id}' },
    create: { method: 'POST',   urlTemplate: '/networkAccess/threatIntelligencePolicies' },
    update: { method: 'PATCH',  urlTemplate: '/networkAccess/threatIntelligencePolicies/{id}' },
    delete: { method: 'DELETE', urlTemplate: '/networkAccess/threatIntelligencePolicies/{id}' },
  },
} as const;
