import { useEffect } from 'react';
import { useIsAuthenticated, useMsal } from '@azure/msal-react';
import { SignInGate } from './ui/shell/SignInGate';
import { GovCloudGuard } from './ui/shell/GovCloudGuard';
import { Workbench } from './ui/shell/Workbench';
import { useDataSource } from './app/dataSourceContext';

export function App() {
  const { sources, activeSourceId, addLiveSource, removeSource, mode } = useDataSource();
  const isAuthenticated = useIsAuthenticated();
  const { accounts } = useMsal();

  // After a successful MSAL sign-in (redirect or silent), register the live
  // source entry if it doesn't exist yet. This covers the case where the user
  // clicked "Sign in" in SignInGate or in the SourceRail — after the MSAL
  // redirect completes, isAuthenticated flips to true and we add the entry.
  useEffect(() => {
    if (isAuthenticated && !sources.some(s => s.type === 'live')) {
      const label = accounts[0]?.username ?? 'Live';
      addLiveSource(label);
    }
  }, [isAuthenticated, sources, addLiveSource, accounts]);

  // Sign-out (or an expired session) must take the live source with it, so no
  // tenant data stays reachable in the UI or the Query cache.
  useEffect(() => {
    if (isAuthenticated) return;
    const live = sources.find(s => s.type === 'live');
    if (live) removeSource(live.id);
  }, [isAuthenticated, sources, removeSource]);

  // No active source → show sign-in / welcome gate.
  if (activeSourceId === null) return <SignInGate />;

  // File mode bypasses the tenant gates.
  if (mode === 'file') return <Workbench />;

  // Live mode requires an authenticated account — never render tenant data
  // based on the source list alone.
  if (!isAuthenticated) return <SignInGate />;

  return (
    <GovCloudGuard>
      <Workbench />
    </GovCloudGuard>
  );
}
