/**
 * Content (File) Policy definition — Tier 2a.
 *
 * Single source of truth for the GSA **Content policy** object type. In the
 * Graph wire format this is `#microsoft.graph.networkaccess.filePolicy`, linked
 * to a profile via a `filePolicyLink`. The admin portal labels it "Content
 * policy"; it provides data-protection / content inspection (the DLP arm of the
 * filtering plane — see the API overview, where `filteringPolicy` "encapsulates
 * … data loss prevention").
 *
 * Shape mirrors the TLS inspection / prompt policies:
 *   • the policy-level default lives in `settings.defaultAction`
 *     ('allow' | 'block'), not a top-level `action`;
 *   • rules (`policyRules`) come from a dedicated sub-route
 *     `/networkAccess/filePolicies/{id}/policyRules` (the profile $expand does
 *     not return them) — same per-policy fetch pattern as TLS/prompt.
 *
 * Microsoft Purview linkage: a content rule references Purview through its
 * **`action`** — `action: "scanPurview"` means the matched content is scanned
 * by Microsoft Purview DLP. There is no per-rule Purview policy id to
 * dereference, so surfacing the Purview signal needs **no extra Graph scope**;
 * the action itself is the correlation.
 *
 * API version: beta
 */

// ---------------------------------------------------------------------------
// Supporting types
// ---------------------------------------------------------------------------

/** Policy-level settings controlling default content-policy behaviour. */
export interface FilePolicySettings {
  /** What happens to traffic not matched by any rule. */
  defaultAction?: 'allow' | 'block';
  [key: string]: unknown;
}

/**
 * One destination a content rule matches. Subtypes seen on the wire:
 *   • `filePolicyFqdnDestination`        — `values` are FQDNs
 *   • `filePolicyWebCategoryDestination` — `values` are web-category names
 * (URL / IP variants follow the same `values[]` shape).
 */
export interface FilePolicyDestination {
  '@odata.type'?: string;
  values?: string[];
  value?: string;
}

/**
 * The file-level matching attributes — what a content rule actually inspects:
 * the traffic direction(s) and the file content types.
 */
export interface FilePolicyFileAttributes {
  /** Comma-separated activity list, e.g. "upload" or "download,upload". */
  activities?: string;
  /** MIME content types the rule matches, e.g. "application/pdf". */
  contentTypes?: string[];
  textContentTypes?: string[];
  [key: string]: unknown;
}

/** The conditions that select traffic/content for a content rule. */
export interface FileRuleMatchingConditions {
  identities?: unknown;
  destinations?: FilePolicyDestination[];
  fileAttributes?: FilePolicyFileAttributes;
  /** e.g. `{ sessionType: "user,agent" }`. */
  sources?: { sessionType?: string; [key: string]: unknown };
  [key: string]: unknown;
}

/**
 * A single content-policy rule (`fileRule`). `action` is PER RULE — observed
 * values include `block` and `scanPurview` (the Microsoft Purview hook).
 */
export interface FileRule {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  priority?: number;
  /** `block` | `scanPurview` | `allow` | … (per-rule). */
  action?: string;
  settings?: { status?: 'enabled' | 'disabled'; [key: string]: unknown };
  matchingConditions?: FileRuleMatchingConditions | null;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A Content (File) Policy — content inspection / data protection. Rules
 * (`policyRules`) are fetched from `/networkAccess/filePolicies/{id}/policyRules`
 * and inlined by the exporter (the profile $expand does not return them).
 */
export interface FilePolicy {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  /** Read-only. Set by the API. */
  version?: string;
  lastModifiedDateTime?: string;
  settings?: FilePolicySettings;
  /** Populated by the exporter from the policy's /policyRules sub-route. */
  policyRules?: FileRule[];
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const FilePolicyDefinition = {
  odataType: '#microsoft.graph.networkaccess.filePolicy',
  displayName: 'Content Policy',

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
    list:   { method: 'GET',    urlTemplate: '/networkAccess/filePolicies' },
    get:    { method: 'GET',    urlTemplate: '/networkAccess/filePolicies/{id}' },
    create: { method: 'POST',   urlTemplate: '/networkAccess/filePolicies' },
    update: { method: 'PATCH',  urlTemplate: '/networkAccess/filePolicies/{id}' },
    delete: { method: 'DELETE', urlTemplate: '/networkAccess/filePolicies/{id}' },
  },
} as const;
