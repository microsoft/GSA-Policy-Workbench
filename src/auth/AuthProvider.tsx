import {
  useCallback,
  useMemo,
  useState,
  type ReactNode,
} from 'react';
import {
  InteractionRequiredAuthError,
} from '@azure/msal-browser';
import { MsalProvider, useMsal } from '@azure/msal-react';
import { useQueryClient } from '@tanstack/react-query';
import { pca, authorityFor, DEFAULT_TENANT } from './msalConfig';
import { CA_DETAIL_SCOPE, OPTIONAL_SCOPES, REQUIRED_SCOPES } from './scopes';
import { setTokenProvider } from '../adapters/graph/client';
import { tracer } from '../app/tracer';
import { AuthContext, type AuthContextValue } from './authContext';

function InnerAuthProvider({ children }: { children: ReactNode }) {
  const { instance, accounts } = useMsal();
  const queryClient = useQueryClient();
  const [tenant, setTenant] = useState<string>(DEFAULT_TENANT);
  const [grantedScopes, setGrantedScopes] = useState<string[]>([]);

  const account = accounts[0] ?? null;

  // Least privilege: only the scope needed to load the policy tree is consented
  // at sign-in. Optional scopes are requested later, on explicit user action.
  const signIn = useCallback(
    async (signInTenant: string) => {
      const result = await instance.loginPopup({
        scopes: [...REQUIRED_SCOPES],
        authority: authorityFor(signInTenant),
      });
      setTenant(signInTenant);
      setGrantedScopes(result.scopes ?? []);
      if (result.account) instance.setActiveAccount(result.account);
    },
    [instance],
  );

  const grantOptionalScopes = useCallback(async () => {
    const result = await instance.acquireTokenPopup({
      scopes: [...OPTIONAL_SCOPES],
      authority: authorityFor(tenant),
    });
    setGrantedScopes((prev) => [...new Set([...prev, ...(result.scopes ?? [])])]);
    if (result.account) instance.setActiveAccount(result.account);
    // Re-fetch so the newly permitted detail is picked up.
    await queryClient.invalidateQueries();
  }, [instance, tenant, queryClient]);

  // Sign-out must leave no tenant data behind on a shared workstation: drop the
  // Query cache and every trace buffer before MSAL clears its own state.
  const signOut = useCallback(async () => {
    setGrantedScopes([]);
    queryClient.clear();
    tracer.clear();
    await instance.logoutPopup();
  }, [instance, queryClient]);

  // Register the token provider the Graph adapter calls. Tries silent first,
  // falls back to interactive ONLY for required scopes. Optional scopes (e.g.
  // Policy.Read.All for CA detail) degrade silently to null so a background
  // load never triggers an unexpected consent popup (spec §2.4).
  setTokenProvider(async (scopes: string[]) => {
    const activeAccount = instance.getActiveAccount() ?? accounts[0];
    if (!activeAccount) return null;
    const isRequired = scopes.some((s) => REQUIRED_SCOPES.includes(s as never));
    try {
      const res = await instance.acquireTokenSilent({
        scopes,
        account: activeAccount,
        authority: authorityFor(tenant),
      });
      return res.accessToken;
    } catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        // Only prompt interactively for a required scope. For an optional
        // scope, degrade silently — the caller handles the missing capability.
        if (!isRequired) return null;
        const res = await instance.acquireTokenPopup({
          scopes,
          authority: authorityFor(tenant),
        });
        return res.accessToken;
      }
      // For a required-scope failure we surface the error; the caller decides.
      if (isRequired) throw err;
      return null;
    }
  });

  const value = useMemo<AuthContextValue>(
    () => ({
      account,
      tenant,
      hasCaDetailScope: grantedScopes.some(
        (s) => s.toLowerCase() === CA_DETAIL_SCOPE.toLowerCase(),
      ),
      hasAllOptionalScopes: OPTIONAL_SCOPES.every((opt) =>
        grantedScopes.some((s) => s.toLowerCase() === opt.toLowerCase()),
      ),
      signIn,
      grantOptionalScopes,
      signOut,
      setTenant,
    }),
    [account, tenant, grantedScopes, signIn, grantOptionalScopes, signOut],
  );

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function AuthProvider({ children }: { children: ReactNode }) {
  return (
    <MsalProvider instance={pca}>
      <InnerAuthProvider>{children}</InnerAuthProvider>
    </MsalProvider>
  );
}
