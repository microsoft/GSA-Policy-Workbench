/**
 * Web Category repository — Tier 2b.
 *
 * Wraps the live web content category checker (loader.fetchWebCategory) in a
 * TanStack Query mutation hook. The UI imports `useWebCategoryLookup` via
 * `query/hooks/` and never calls Graph directly.
 *
 * Live-only: the checker has no fixture, so it is unavailable in file mode
 * (the caller should disable it there). Uses the already-required
 * `NetworkAccess.Read.All` scope — no new consent.
 */

import { useMutation } from '@tanstack/react-query';
import { fetchWebCategory } from '../../adapters/graph/loader';
import type { WebCategory } from '../definitions/WebContentFilteringPolicy.definition';

/**
 * On-demand web content category lookup for a URL/host.
 *
 * @example
 * const lookup = useWebCategoryLookup();
 * lookup.mutate('msn.com/en-us/sports');
 * // lookup.data → { displayName: 'Sports', group: 'GeneralSurfing', … }
 */
export function useWebCategoryLookup() {
  return useMutation<WebCategory, Error, string>({
    mutationFn: (url) => fetchWebCategory(url),
  });
}
