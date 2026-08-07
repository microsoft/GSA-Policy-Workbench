/**
 * TanStack Query hook for the Security Profile tree.
 *
 * Re-exports `useProfileTree` from the SecurityProfile repository so that
 * UI components import from `query/hooks/` (Tier 1 boundary) rather than
 * reaching into `model/repositories/` directly.
 */

export { useProfileTree } from '../../model/repositories/SecurityProfile.repository';

