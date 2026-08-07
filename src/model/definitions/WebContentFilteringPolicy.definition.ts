/**
 * Web Content Filtering Policy definition — Tier 2a.
 *
 * Single source of truth for the Web Content Filtering Policy GSA object type.
 * In the Graph API this is a concrete subtype of `filteringPolicy` (abstract),
 * distinguished by @odata.type.
 *
 * A filtering policy is a container for filtering rules. It is linked to one
 * or more Security Profiles via filteringPolicyLink objects.
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-filteringpolicy
 * API version: beta
 */

import type { FilteringRule } from './FilteringRule.definition';

// ---------------------------------------------------------------------------
// Supporting types — the policy link (join between profile and policy)
// ---------------------------------------------------------------------------

/**
 * The join object between a SecurityProfile and a FilteringPolicy.
 * Carries per-profile settings (action, priority, state, loggingState).
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-filteringpolicylink
 */
export interface FilteringPolicyLink {
  '@odata.type'?: string;
  id: string;
  /** Whether the link is active. */
  state?: 'enabled' | 'disabled';
  /** Version string from the API. */
  version?: string;
  /** Default action when no rule matches — applied at the profile+policy level. */
  action?: 'block' | 'allow';
  /** Whether logging is active for this link. */
  loggingState?: 'enabled' | 'disabled';
  /** Processing priority of this policy within the profile (lower = higher priority). */
  priority?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  /** Expanded via $expand=policies($expand=policy). Always present when expanded. */
  policy?: WebContentFilteringPolicy;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A Web Content Filtering Policy — a named set of filtering rules.
 * Contains policyRules (FilteringRule[]) when expanded.
 *
 * The `@odata.type` discriminant distinguishes it from TLS / CloudFirewall /
 * ThreatIntelligence policies when multiple types appear in the same list.
 */
export interface WebContentFilteringPolicy {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  version?: string;
  /**
   * Default action when no rule matches. In live Graph this lives on the
   * policy; some fixtures carry it on the link instead (see FilteringPolicyLink).
   */
  action?: 'block' | 'allow';
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  /** Populated when expanded via $expand=policyRules. */
  policyRules?: FilteringRule[];
}

// ---------------------------------------------------------------------------
// Web category lookup
// ---------------------------------------------------------------------------

/**
 * Result of the web content category checker — the category a given URL/host
 * resolves to. Returned by the beta `getWebCategoriesByUrl` function (preview).
 *
 * Graph docs: https://learn.microsoft.com/en-us/entra/global-secure-access/how-to-check-web-content-filtering-categories
 */
export interface WebCategory {
  '@odata.type'?: string;
  /** Internal category name (e.g. "Sports"). */
  name?: string;
  /** Friendly category name shown to the admin. */
  displayName?: string;
  /** Parent category group (e.g. "GeneralSurfing"). */
  group?: string;
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const WebContentFilteringPolicyDefinition = {
  odataType: '#microsoft.graph.networkaccess.webContentFilteringPolicy',
  displayName: 'Web Content Filtering Policy',

  properties: {
    id:                   { label: 'ID' },
    name:                 { label: 'Policy Name' },
    description:          { label: 'Description' },
    version:              { label: 'Version' },
    action:               { label: 'Action' },
    createdDateTime:      { label: 'Created' },
    lastModifiedDateTime: { label: 'Last Modified' },
    policyRules:          { label: 'Rules' },
  },

  operations: {
    list:   { method: 'GET',    urlTemplate: '/networkAccess/filteringPolicies' },
    get:    { method: 'GET',    urlTemplate: '/networkAccess/filteringPolicies/{id}' },
    create: { method: 'POST',   urlTemplate: '/networkAccess/filteringPolicies' },
    update: { method: 'PATCH',  urlTemplate: '/networkAccess/filteringPolicies/{id}' },
    delete: { method: 'DELETE', urlTemplate: '/networkAccess/filteringPolicies/{id}' },
    /**
     * Web category checker (preview, beta) — resolve which content category a
     * URL/host belongs to. The `url` is passed as the `@url` function parameter;
     * the audit endpoint stays templated (the queried host is never logged).
     */
    checkWebCategory: {
      method: 'GET',
      urlTemplate:
        "/networkAccess/connectivity/microsoft.graph.networkaccess.getWebCategoriesByUrl(url='@url')",
    },
  },
} as const;
