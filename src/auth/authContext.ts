import { createContext, useContext } from 'react';
import type { AccountInfo } from '@azure/msal-browser';

export interface AuthContextValue {
  account: AccountInfo | null;
  tenant: string;
  /** True when the optional Policy.Read.All scope was granted. */
  hasCaDetailScope: boolean;
  /** True when every optional scope has been granted. */
  hasAllOptionalScopes: boolean;
  signIn: (tenant: string) => Promise<void>;
  /** Interactive consent for the optional read scopes. User-initiated only. */
  grantOptionalScopes: () => Promise<void>;
  signOut: () => Promise<void>;
  setTenant: (tenant: string) => void;
}

export const AuthContext = createContext<AuthContextValue | undefined>(
  undefined,
);

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used within an AuthProvider');
  return ctx;
}
