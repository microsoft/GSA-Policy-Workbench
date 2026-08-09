/**
 * TanStack Query hook for the web content category checker.
 *
 * Re-exports `useWebCategoryLookup` from the WebCategory repository so that UI
 * components import from `query/hooks/` (Tier 1 boundary) rather than reaching
 * into `model/repositories/` directly.
 */

export { useWebCategoryLookup } from '../../model/repositories/WebCategory.repository';
