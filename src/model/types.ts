/**
 * Shared model types used across definition files and repositories.
 *
 * Keep this file minimal — object-specific types belong in their own
 * definition file. Only cross-cutting types belong here.
 */

// ---------------------------------------------------------------------------
// Graph wire-format utilities
// ---------------------------------------------------------------------------

/**
 * The standard OData collection wrapper returned by Graph list endpoints.
 * `@odata.nextLink` is present when more pages are available.
 */
export interface ODataCollection<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

// ---------------------------------------------------------------------------
// Cross-cutting domain types
// ---------------------------------------------------------------------------

/**
 * The status field used across multiple GSA resource types.
 * Maps to `microsoft.graph.networkaccess.status`.
 */
export type GsaStatus = 'enabled' | 'disabled';

/**
 * A discriminated union of all filtering policy types that can appear in a
 * Security Profile's `policies` relationship. The `@odata.type` field on
 * the nested `policy` object identifies the concrete type.
 *
 * New policy types are added here when their definition file is implemented.
 */
export type KnownPolicyOdataType =
  | '#microsoft.graph.networkaccess.webContentFilteringPolicy'
  | '#microsoft.graph.networkaccess.tlsInspectionPolicy'
  | '#microsoft.graph.networkaccess.cloudFirewallPolicy'
  | '#microsoft.graph.networkaccess.threatIntelligencePolicy';

