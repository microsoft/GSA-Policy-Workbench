/**
 * Cloud Firewall Policy definition — Tier 2a.
 *
 * Single source of truth for the Cloud Firewall Policy GSA object type.
 * Provides Layer 3/4 (network) protection by monitoring and controlling
 * network traffic via the Global Secure Access service.
 *
 * Shape mirrors the TLS inspection / prompt / content policies:
 *   • the policy-level default lives in `settings.defaultAction`
 *     ('allow' | 'block'), not a top-level `action`;
 *   • rules (`policyRules`) come from a dedicated sub-route
 *     `/networkAccess/cloudFirewallPolicies/{id}/policyRules` (the profile
 *     $expand does not return them) — same per-policy fetch pattern as the
 *     other rule-bearing policy kinds.
 *
 * A `cloudFirewallRule` matches on 5-tuple-style conditions: source and
 * destination IP addresses / FQDNs, ports, and transport protocol. The
 * per-rule `action` is `allow` | `block`.
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-cloudfirewallpolicy
 * API version: beta
 */

// ---------------------------------------------------------------------------
// Supporting types
// ---------------------------------------------------------------------------

/**
 * Policy-level settings for the Cloud Firewall, including the default action
 * applied when no rule matches.
 */
export interface CloudFirewallPolicySettings {
  /** The action taken when no rule matches the traffic. */
  defaultAction?: 'allow' | 'block';
  [key: string]: unknown;
}

/**
 * One source/destination address group on a cloud firewall rule. Subtypes seen
 * on the wire:
 *   • `cloudFirewallSourceIpAddress`        — `values` are source IPs / CIDRs
 *   • `cloudFirewallDestinationIpAddress`   — `values` are destination IPs / CIDRs
 *   • `cloudFirewallDestinationFqdnAddress` — `values` are destination FQDNs
 */
export interface CloudFirewallAddress {
  '@odata.type'?: string;
  values?: string[];
}

/** The source half of a rule's matching conditions (ports + addresses). */
export interface CloudFirewallSource {
  /** Port ranges, e.g. ["0-65535"]. */
  ports?: string[];
  addresses?: CloudFirewallAddress[];
}

/** The destination half of a rule's matching conditions (ports + protocol + addresses). */
export interface CloudFirewallDestination {
  /** Port ranges / single ports, e.g. ["88", "445"]. */
  ports?: string[];
  /** Transport protocol(s), e.g. "tcp", "udp", or "tcp,udp". */
  protocols?: string;
  addresses?: CloudFirewallAddress[];
}

/** The conditions that select traffic for a cloud firewall rule. */
export interface CloudFirewallRuleMatchingConditions {
  sources?: CloudFirewallSource;
  destinations?: CloudFirewallDestination;
  [key: string]: unknown;
}

/**
 * A single Cloud Firewall rule (`cloudFirewallRule`). `action` is PER RULE
 * (`allow` | `block`); `settings.status` carries the rule's enabled state.
 */
export interface CloudFirewallRule {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  priority?: number;
  /** `allow` | `block` (per-rule). */
  action?: 'allow' | 'block';
  settings?: { status?: 'enabled' | 'disabled'; [key: string]: unknown };
  matchingConditions?: CloudFirewallRuleMatchingConditions | null;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A Cloud Firewall Policy — Layer 3/4 firewall for network traffic. Rules
 * (`cloudFirewallRule`) are fetched from
 * `/networkAccess/cloudFirewallPolicies/{id}/policyRules` and inlined by the
 * exporter (the profile $expand does not return them).
 */
export interface CloudFirewallPolicy {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  /** Read-only. Set by the API. */
  version?: string;
  lastModifiedDateTime?: string;
  /** Required. Default action + other firewall settings. */
  settings?: CloudFirewallPolicySettings;
  /** Populated when expanded via $expand=policyRules. */
  policyRules?: CloudFirewallRule[];
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const CloudFirewallPolicyDefinition = {
  odataType: '#microsoft.graph.networkaccess.cloudFirewallPolicy',
  displayName: 'Cloud Firewall Policy',

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
    list:   { method: 'GET',    urlTemplate: '/networkAccess/cloudFirewallPolicies' },
    get:    { method: 'GET',    urlTemplate: '/networkAccess/cloudFirewallPolicies/{id}' },
    create: { method: 'POST',   urlTemplate: '/networkAccess/cloudFirewallPolicies' },
    update: { method: 'PATCH',  urlTemplate: '/networkAccess/cloudFirewallPolicies/{id}' },
    delete: { method: 'DELETE', urlTemplate: '/networkAccess/cloudFirewallPolicies/{id}' },
  },
} as const;
