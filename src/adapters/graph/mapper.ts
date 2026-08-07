/**
 * Generic Graph Mapper — Tier 3.
 *
 * Converts raw Graph wire-format JSON to typed domain objects by picking only
 * the fields declared in a definition's `properties` map, plus the `id` field
 * and the `@odata.type` discriminant.
 *
 * This is intentionally not a deep validator — unknown Graph fields are
 * silently discarded, and missing optional fields are omitted (not defaulted).
 * This keeps the mapper simple and tolerant of beta schema drift.
 *
 * See docs/architecture.md §6.2 for the mapper contract.
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** The shape of a definition's `properties` map (only what the mapper needs). */
type PropertyMap = Record<string, { label: string }>;

// ---------------------------------------------------------------------------
// Core mapping functions
// ---------------------------------------------------------------------------

/**
 * Map a single raw Graph wire object to a typed domain object.
 *
 * Only fields listed in `properties` are copied to the output. The `id`
 * field and `@odata.type` are always included when present.
 *
 * Throws if the wire object is not an object or is missing an `id` string —
 * this represents a malformed Graph response, not a user error.
 */
export function mapOne<T extends { id: string }>(
  wireItem: unknown,
  properties: PropertyMap,
): T {
  if (wireItem === null || typeof wireItem !== 'object') {
    throw new Error('mapOne: wireItem is not an object.');
  }

  const src = wireItem as Record<string, unknown>;

  if (typeof src['id'] !== 'string' || src['id'] === '') {
    throw new Error('mapOne: wireItem is missing a non-empty "id" string field.');
  }

  const result: Record<string, unknown> = {};

  // Always carry the @odata.type discriminant for polymorphic lists.
  if (src['@odata.type'] !== undefined) {
    result['@odata.type'] = src['@odata.type'];
  }

  // Copy only declared properties — discard everything else.
  for (const key of Object.keys(properties)) {
    if (src[key] !== undefined) {
      result[key] = src[key];
    }
  }

  return result as T;
}

/**
 * Map an array of raw Graph wire objects to typed domain objects.
 *
 * Items that fail mapping (e.g., missing `id`) are silently dropped so that
 * one malformed item cannot break the entire list. This matches the beta API's
 * known tendency to return partially-formed objects during schema transitions.
 */
export function mapCollection<T extends { id: string }>(
  wireItems: unknown[],
  properties: PropertyMap,
): T[] {
  const results: T[] = [];
  for (const item of wireItems) {
    try {
      results.push(mapOne<T>(item, properties));
    } catch {
      // Silently drop — beta schema drift should not crash the UI.
    }
  }
  return results;
}

// ---------------------------------------------------------------------------
// Convenience: list from a full OData collection response
// ---------------------------------------------------------------------------

/**
 * Extract and map items from a raw OData collection response body.
 * Equivalent to `mapCollection(body.value ?? [], properties)`.
 */
export function mapODataCollection<T extends { id: string }>(
  body: unknown,
  properties: PropertyMap,
): T[] {
  if (body === null || typeof body !== 'object') return [];
  const collection = body as Record<string, unknown>;
  const value = collection['value'];
  return mapCollection<T>(Array.isArray(value) ? value : [], properties);
}

