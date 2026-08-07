/**
 * Generic Graph Connector — Tier 3.
 *
 * Translates a URL template + operation name into a real Graph HTTP call.
 * Injects auth tokens (via graphGet/graphBatch in client.ts), handles
 * pagination, and routes every call through the audit interceptor.
 *
 * Rules (from architecture.md §6):
 * - Never contains object-specific logic.
 * - URL templates (not populated URLs) are passed to the interceptor.
 * - Adding a new object type = adding a definition file, NOT changing this file.
 *
 * See docs/architecture.md §6 for the full contract.
 */

import { graphGet } from './client';
import { REQUIRED_SCOPES } from '../../auth/scopes';
import type { ODataCollection } from '../../model/types';

// ---------------------------------------------------------------------------
// Options
// ---------------------------------------------------------------------------

export interface ConnectorOptions {
  /**
   * URL template parameters to substitute, e.g. `{ id: '…', policyId: '…' }`.
   * Values are encoded with `encodeURIComponent` before substitution.
   */
  params?: Record<string, string>;
  /**
   * OData query string to append to the URL (without the leading `?`).
   * Example: `'$expand=policyRules&$select=id,name'`
   */
  query?: string;
  /**
   * Auth scopes. Defaults to `REQUIRED_SCOPES` (`NetworkAccess.Read.All`).
   * Pass `[CA_DETAIL_SCOPE]` for Conditional Access endpoints.
   */
  scopes?: readonly string[];
  /** Logical feature name recorded by the audit interceptor. */
  feature: string;
  signal?: AbortSignal;
}

// ---------------------------------------------------------------------------
// URL template resolution
// ---------------------------------------------------------------------------

/**
 * Replace `{placeholder}` tokens in a URL template with URL-encoded values.
 *
 * Throws if a placeholder has no corresponding value — this is a programmer
 * error, not a runtime condition, so throwing is correct.
 *
 * The populated URL is used for the actual HTTP call. The original template
 * is passed to the interceptor so that no IDs or names appear in audit records.
 */
export function resolveUrlTemplate(
  template: string,
  params: Record<string, string> = {},
): string {
  return template.replace(/\{(\w+)\}/g, (_, key: string) => {
    const val = params[key];
    if (val === undefined || val === '') {
      throw new Error(
        `resolveUrlTemplate: missing required parameter {${key}} in template "${template}".`,
      );
    }
    return encodeURIComponent(val);
  });
}

// ---------------------------------------------------------------------------
// Public API — V0 read operations
// ---------------------------------------------------------------------------

/**
 * Fetch a paginated list of objects from a Graph collection endpoint.
 * Follows `@odata.nextLink` automatically until all pages are consumed.
 *
 * The `urlTemplate` (with `{placeholder}` tokens) is recorded by the
 * interceptor; the populated URL is used only for the HTTP call.
 *
 * @example
 * const profiles = await connectorList<SecurityProfile>(
 *   SecurityProfileDefinition.operations.list.urlTemplate,
 *   { feature: 'load-table', query: '$expand=policies' },
 * );
 */
export async function connectorList<T>(
  urlTemplate: string,
  opts: ConnectorOptions,
): Promise<T[]> {
  const path = resolveUrlTemplate(urlTemplate, opts.params);
  const url = opts.query ? `${path}?${opts.query}` : path;
  const scopes = [...(opts.scopes ?? REQUIRED_SCOPES)];

  const result = await graphGet<ODataCollection<T>>(url, {
    scopes,
    feature: opts.feature,
    endpoint: urlTemplate,  // template, not populated URL — privacy rule §6.3
    paginate: true,
    signal: opts.signal,
  });

  return result.value;
}

/**
 * Fetch a single object from a Graph resource endpoint (no pagination).
 *
 * @example
 * const profile = await connectorGet<SecurityProfile>(
 *   SecurityProfileDefinition.operations.get.urlTemplate,
 *   { params: { id: profileId }, feature: 'profile-detail' },
 * );
 */
export async function connectorGet<T>(
  urlTemplate: string,
  opts: ConnectorOptions,
): Promise<T> {
  const path = resolveUrlTemplate(urlTemplate, opts.params);
  const url = opts.query ? `${path}?${opts.query}` : path;
  const scopes = [...(opts.scopes ?? REQUIRED_SCOPES)];

  return graphGet<T>(url, {
    scopes,
    feature: opts.feature,
    endpoint: urlTemplate,  // template, not populated URL — privacy rule §6.3
    signal: opts.signal,
  });
}

// ---------------------------------------------------------------------------
// V1+ write operations — placeholders (not yet active in V0)
// ---------------------------------------------------------------------------

/**
 * Issue a PATCH mutation. Not active in V0 — all write operations are blocked
 * at the Repository layer until V1+.
 *
 * See docs/architecture.md §7 and docs/spec.md §5 for the write contract.
 */
export async function connectorUpdate<TBody, TResult>(
  urlTemplate: string,
  body: TBody,
  opts: ConnectorOptions,
): Promise<TResult> {
  // V1+ implementation goes here.
  // Must: block UI during in-flight, show error inline on failure,
  // never use optimistic updates (confirmed server state only).
  void body;
  void opts;
  throw new Error(`connectorUpdate is not yet implemented (V1+). Template: ${urlTemplate}`);
}

/**
 * Issue a POST to create a new resource. Not active in V0.
 */
export async function connectorCreate<TBody, TResult>(
  urlTemplate: string,
  body: TBody,
  opts: ConnectorOptions,
): Promise<TResult> {
  void body;
  void opts;
  throw new Error(`connectorCreate is not yet implemented (V1+). Template: ${urlTemplate}`);
}

/**
 * Issue a DELETE. Not active in V0.
 */
export async function connectorDelete(
  urlTemplate: string,
  opts: ConnectorOptions,
): Promise<void> {
  void opts;
  throw new Error(`connectorDelete is not yet implemented (V1+). Template: ${urlTemplate}`);
}

