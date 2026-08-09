/**
 * Filtering Rule definition — Tier 2a.
 *
 * Single source of truth for the FilteringRule GSA object type.
 * Contains the domain type, supporting types, display labels, and
 * Graph operation URL templates.
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-filteringrule
 * API version: beta
 */

// ---------------------------------------------------------------------------
// Supporting types — ruleDestination subtypes
// ---------------------------------------------------------------------------

/** Destination: a web category (e.g. "Gambling", "Adult Content"). */
export interface WebCategoryDestination {
  '@odata.type': '#microsoft.graph.networkaccess.webCategory';
  /** Unique identifier for the category (used as key in rules). */
  name: string;
  /** Human-readable label for display. */
  displayName?: string;
  /** Grouping label (e.g. "Adult"). */
  group?: string;
}

/** Destination: a fully-qualified domain name. */
export interface FqdnDestination {
  '@odata.type': '#microsoft.graph.networkaccess.fqdn';
  value: string;
}

/** Destination: a URL (protocol + domain + path). */
export interface UrlDestination {
  '@odata.type': '#microsoft.graph.networkaccess.url';
  value: string;
}

/** Destination: a single IPv4/IPv6 address. */
export interface IpAddressDestination {
  '@odata.type': '#microsoft.graph.networkaccess.ipAddress';
  value: string;
}

/** Destination: an IP range (start–end). */
export interface IpRangeDestination {
  '@odata.type': '#microsoft.graph.networkaccess.ipRange';
  /** Start of the range, e.g. "10.0.0.1". */
  beginAddress: string;
  /** End of the range, e.g. "10.0.0.254". */
  endAddress: string;
}

/** Destination: a CIDR subnet, e.g. "10.0.0.0/24". */
export interface IpSubnetDestination {
  '@odata.type': '#microsoft.graph.networkaccess.ipSubnet';
  value: string;
}

/** Unknown future destination type — preserved for forward compatibility. */
export interface UnknownDestination {
  '@odata.type': string;
  [key: string]: unknown;
}

/** Discriminated union of all known ruleDestination subtypes. */
export type RuleDestination =
  | WebCategoryDestination
  | FqdnDestination
  | UrlDestination
  | IpAddressDestination
  | IpRangeDestination
  | IpSubnetDestination
  | UnknownDestination;

/**
 * The network destination type. Determines the shape of each destination
 * in the `destinations` array and controls UI rendering.
 */
export type NetworkDestinationType =
  | 'url'
  | 'fqdn'
  | 'webCategory'
  | 'ipAddress'
  | 'ipRange'
  | 'ipSubnet'
  | 'unknownFutureValue';

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A filtering rule — the smallest unit of policy. Lives inside a
 * filteringPolicy (via policyRules). The `@odata.type` discriminant
 * identifies the concrete subtype (fqdnFilteringRule, urlDestinationFilteringRule,
 * webCategoryFilteringRule).
 */
export interface FilteringRule {
  '@odata.type'?: string;
  id: string;
  name?: string;
  /** Processing order within the policy (lower = evaluated first). */
  priority?: number;
  ruleType?: NetworkDestinationType;
  destinations?: RuleDestination[];
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

/**
 * Definition constant — read by the connector (operations → URLs),
 * the mapper (properties → display labels), and the UI (column headers).
 */
export const FilteringRuleDefinition = {
  odataType: '#microsoft.graph.networkaccess.filteringRule',
  displayName: 'Filtering Rule',

  /**
   * Property metadata — one entry per domain field.
   * `label` is the column header / form label shown in the UI.
   */
  properties: {
    id:           { label: 'ID' },
    name:         { label: 'Rule Name' },
    priority:     { label: 'Priority' },
    ruleType:     { label: 'Rule Type' },
    destinations: { label: 'Destinations' },
  },

  /**
   * Graph operation URL templates — {policyId} and {id} are replaced at
   * runtime by the connector. These are the template strings recorded in the
   * audit interceptor (never populated URLs).
   */
  operations: {
    list:   { method: 'GET',    urlTemplate: '/networkAccess/filteringPolicies/{policyId}/policyRules' },
    get:    { method: 'GET',    urlTemplate: '/networkAccess/filteringPolicies/{policyId}/policyRules/{id}' },
    create: { method: 'POST',   urlTemplate: '/networkAccess/filteringPolicies/{policyId}/policyRules' },
    update: { method: 'PATCH',  urlTemplate: '/networkAccess/filteringPolicies/{policyId}/policyRules/{id}' },
    delete: { method: 'DELETE', urlTemplate: '/networkAccess/filteringPolicies/{policyId}/policyRules/{id}' },
  },
} as const;
