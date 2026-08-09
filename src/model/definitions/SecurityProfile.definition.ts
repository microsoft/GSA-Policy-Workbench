/**
 * Security Profile definition — Tier 2a.
 *
 * Single source of truth for the Security Profile (filteringProfile) GSA
 * object type. A Security Profile is the top-level container that groups
 * filtering policies and is linked to Conditional Access policies to scope
 * which users / devices it applies to.
 *
 * Graph resource name: filteringProfile
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-filteringprofile
 * API version: beta
 */

import type { FilteringPolicyLink } from './WebContentFilteringPolicy.definition';

// ---------------------------------------------------------------------------
// Supporting types
// ---------------------------------------------------------------------------

/**
 * A stub of a Conditional Access policy as projected onto a Security Profile.
 * Full CA policy detail (conditions, grant controls) requires Policy.Read.All
 * and is loaded separately by the ConditionalAccess repository.
 */
export interface LinkedCaPolicy {
  id: string;
  displayName?: string;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A Security Profile (filteringProfile) — the top-level container.
 *
 * In the Graph API this is `networkaccess.filteringProfile`. It:
 * - Links to one or more filtering policies via `policies` (FilteringPolicyLink[])
 * - Links to Conditional Access policies via `conditionalAccessPolicies`
 * - Has a `priority` that determines processing order when multiple profiles match
 *
 * Relationships are populated when expanded:
 *   $expand=policies($expand=policy($expand=policyRules)),conditionalAccessPolicies
 */
export interface SecurityProfile {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  state?: 'enabled' | 'disabled';
  /** Lower number = higher priority. */
  priority?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  /** Expanded: the linked filtering policies with their rules. */
  policies?: FilteringPolicyLink[];
  /** Always present: CA policy stubs linked to this profile. */
  conditionalAccessPolicies?: LinkedCaPolicy[];
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const SecurityProfileDefinition = {
  odataType: '#microsoft.graph.networkaccess.filteringProfile',
  displayName: 'Security Profile',

  properties: {
    id:                        { label: 'ID' },
    name:                      { label: 'Profile Name' },
    description:               { label: 'Description' },
    state:                     { label: 'State' },
    priority:                  { label: 'Priority' },
    createdDateTime:           { label: 'Created' },
    lastModifiedDateTime:      { label: 'Last Modified' },
    policies:                  { label: 'Linked Policies' },
    conditionalAccessPolicies: { label: 'Conditional Access Policies' },
  },

  operations: {
    list:   { method: 'GET',   urlTemplate: '/networkAccess/filteringProfiles' },
    get:    { method: 'GET',   urlTemplate: '/networkAccess/filteringProfiles/{id}' },
    // Note: Create is not documented in beta — profiles appear to be managed
    // through the Entra portal or via policy links. Included for future use.
    update: { method: 'PATCH', urlTemplate: '/networkAccess/filteringProfiles/{id}' },
  },
} as const;
