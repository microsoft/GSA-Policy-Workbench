/**
 * Conditional Access repository — Tier 2b.
 *
 * Fetches full Conditional Access policy detail for the where-used drawer.
 * Requires the optional `Policy.Read.All` scope. When the scope is absent
 * (`hasCaDetailScope === false`), the hook is disabled and the drawer falls
 * back to the stub (id + displayName) already present on the Security Profile.
 */

import { useQuery } from '@tanstack/react-query';
import { graphBatch } from '../../adapters/graph/client';
import { CA_DETAIL_SCOPE } from '../../auth/scopes';
import type { ConditionalAccessPolicy } from '../definitions/ConditionalAccessPolicy.definition';
import { queryKeys } from '../../query/keys';

const FEATURE = 'where-used';

/**
 * Fetch and cache full Conditional Access policy detail for a set of ids.
 *
 * @param ids     - The CA policy ids to fetch (from SecurityProfile.conditionalAccessPolicies)
 * @param enabled - Set to `hasCaDetailScope` from useAuth(). Disables the query
 *                  when the Policy.Read.All scope was not granted.
 *
 * Returns a `Map<id, ConditionalAccessPolicy>` — callers look up by id to
 * display conditions, grant controls, etc. in the where-used drawer.
 *
 * @example
 * const { hasCaDetailScope } = useAuth();
 * const caIds = profile.conditionalAccessPolicies?.map(p => p.id) ?? [];
 * const { data: caMap } = useCaPolicies(caIds, hasCaDetailScope);
 */
export function useCaPolicies(
  ids: readonly string[],
  enabled: boolean,
) {
  return useQuery<Map<string, ConditionalAccessPolicy>, Error>({
    queryKey: queryKeys.conditionalAccessPolicies(ids),
    queryFn: () => fetchCaPolicies([...ids]),
    enabled: enabled && ids.length > 0,
    staleTime: 5 * 60_000, // CA policies change rarely — cache for 5 min
  });
}

// ---------------------------------------------------------------------------
// Private fetch function
// ---------------------------------------------------------------------------

async function fetchCaPolicies(
  ids: string[],
): Promise<Map<string, ConditionalAccessPolicy>> {
  const result = new Map<string, ConditionalAccessPolicy>();
  if (ids.length === 0) return result;

  const requests = ids.map((id) => ({
    id,
    method: 'GET' as const,
    // Note: populated URL inside $batch body is acceptable — the batch
    // envelope endpoint (/$batch) is what the interceptor records.
    url: `/identity/conditionalAccess/policies/${id}`,
  }));

  const responses = await graphBatch<ConditionalAccessPolicy>(requests, {
    scopes: [CA_DETAIL_SCOPE],
    feature: FEATURE,
  });

  for (const [id, resp] of responses) {
    if (resp.status >= 200 && resp.status < 300 && resp.body?.id) {
      result.set(id, resp.body);
    }
  }

  return result;
}

