/**
 * Security Profile repository — Tier 2b.
 *
 * Wraps the adaptive profile tree loader in a TanStack Query hook.
 * Exposes `useProfileTree()` to UI components.
 *
 * What this does NOT do:
 * - It does not call Graph directly (that is loader.ts / connector.ts).
 * - It does not transform data (the loader returns typed domain objects).
 * - It does not store anything — TanStack Query caches in browser RAM.
 *
 * See docs/architecture.md §5 for the repository contract.
 */

import { useQuery } from '@tanstack/react-query';
import { loadProfileTree } from '../../adapters/graph/loader';
import type { ProfileTreeResult } from '../../adapters/graph/loader';
import { queryKeys } from '../../query/keys';

/**
 * Fetch and cache the full Security Profile tree:
 *   SecurityProfile[] → FilteringPolicyLink[] → WebContentFilteringPolicy → FilteringRule[]
 *
 * Uses the adaptive Branch A / Branch B loading strategy (see loader.ts).
 * Returns the same `{ data, isLoading, isError, error }` shape as all
 * TanStack Query hooks.
 *
 * @param sourceId  Active data-source id (from DataSourceContext). Included in
 *                  the query key so each source has its own cache slot — switching
 *                  sources preserves the previously fetched data for each source.
 *
 * @example
 * const { data, isLoading, isError } = useProfileTree(activeSourceId);
 * const profiles = data?.profiles ?? [];
 */
export function useProfileTree(sourceId: string | null) {
  return useQuery<ProfileTreeResult, Error>({
    queryKey: queryKeys.securityProfiles(sourceId ?? '__disabled__'),
    queryFn: ({ signal }) => loadProfileTree(signal),
    staleTime: 60_000,
    enabled: sourceId != null,
  });
}


