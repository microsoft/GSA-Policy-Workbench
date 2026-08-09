import { useMutation } from '@tanstack/react-query';
import { resolveUserByUpn, type ResolvedUser } from '../../adapters/graph/loader';

export type { ResolvedUser };

export function useResolveUser() {
  return useMutation({ mutationFn: (upn: string) => resolveUserByUpn(upn) });
}
