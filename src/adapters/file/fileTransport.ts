/**
 * File transport — file-based data source (see docs/architecture.md §6.4).
 *
 * Returns a `GraphFetch` that answers Graph round-trips from a parsed fixture
 * instead of the network. Installed via `setGraphTransport()` when entering
 * file mode. Because it returns real `Response` objects, the connector, loader,
 * mapper, and audit interceptor in Tier 3 run completely unchanged.
 *
 * Status: STUB. The Branch A path (expanded profile tree) is implemented.
 * Branch B (`$batch` per-policy `policyRules`) and CA-detail endpoints are
 * TODOs pending the open decisions in §6.4.
 */

import type { GraphFetch } from '../graph/transport';
import type { FixtureDoc } from './fixture';

/** Build a synthetic JSON `Response`, mirroring what `fetch` would return. */
function jsonResponse(body: unknown, status: number): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

/**
 * Create a file-backed transport bound to a parsed fixture document.
 *
 * Matching is done on the request path (the populated URL minus origin):
 * - `…/networkAccess/filteringProfiles…`  → the fixture's profile tree (Branch A)
 * - anything else                          → 404 with a clear diagnostic body
 */
export function createFileTransport(doc: FixtureDoc): GraphFetch {
  return (url) => {
    const path = url.replace(/^https?:\/\/[^/]+/, '');

    if (path.includes('/networkAccess/filteringProfiles')) {
      return Promise.resolve(jsonResponse(doc.filteringProfiles, 200));
    }

    // TODO(file-mode §6.4): answer `/$batch` policyRules (Branch B) and
    // Conditional Access detail endpoints once the fixture envelope is decided.
    return Promise.resolve(
      jsonResponse(
        {
          error: {
            code: 'fileMode_unmapped',
            message: `No fixture mapping for "${path}". ` +
              'Provide policyRules inline (Branch A) or extend createFileTransport.',
          },
        },
        404,
      ),
    );
  };
}
