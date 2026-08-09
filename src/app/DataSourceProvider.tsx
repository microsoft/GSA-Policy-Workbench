/**
 * DataSourceProvider — Tier 1 (file-based data source).
 *
 * Manages a list of named data sources (at most one live Graph connection and
 * any number of loaded fixture files). Switching the active source updates the
 * pluggable transport seam and clears the TanStack Query cache so the Workbench
 * re-fetches from the new source immediately — no page reload required.
 *
 * On mount the provider checks whether MSAL already has an active account (i.e.
 * the user was signed in during a previous session) and, if so, pre-populates a
 * live source so the Workbench renders without a flash of the sign-in gate.
 */

import { useCallback, useMemo, useRef, useState, type ReactNode } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { setAuthBypass } from '../adapters/graph/client';
import { setGraphTransport } from '../adapters/graph/transport';
import { createFileTransport } from '../adapters/file/fileTransport';
import { parseFixture, type FixtureDoc } from '../adapters/file/fixture';
import { pca } from '../auth/msalConfig';
import {
  DataSourceContext,
  type DataSourceContextValue,
  type DataSourceMode,
  type SourceEntry,
} from './dataSourceContext';

// ── helpers ─────────────────────────────────────────────────────────────────

let _seq = 1;
function nextFileId() { return `file-${_seq++}`; }
const LIVE_ID = 'live';

function initialSources(): { sources: SourceEntry[]; activeId: string | null } {
  const account = pca.getActiveAccount() ?? pca.getAllAccounts()[0] ?? null;
  if (account) {
    return {
      sources: [{ id: LIVE_ID, type: 'live', label: account.username }],
      activeId: LIVE_ID,
    };
  }
  return { sources: [], activeId: null };
}

// ── provider ─────────────────────────────────────────────────────────────────

export function DataSourceProvider({ children }: { children: ReactNode }) {
  const queryClient = useQueryClient();

  const init = useMemo(initialSources, []);
  const [sources, setSources] = useState<SourceEntry[]>(init.sources);
  const [activeSourceId, setActiveSourceId] = useState<string | null>(init.activeId);
  const [sourceError, setSourceError] = useState<string | null>(null);

  // Fixture documents keyed by source id — stored in a ref so swapping them
  // does not cause extra renders.
  const docMap = useRef<Map<string, FixtureDoc>>(new Map());

  // ── internal helper ───────────────────────────────────────────────────────

  /**
   * Point the transport seam at `entry`. Returns an error message when the
   * source cannot be activated, or null on success.
   *
   * Fails closed: the auth bypass is only ever enabled together with a file
   * transport. Leaving `authBypassed` true with no transport installed would
   * send the next Graph call to the live service carrying the literal
   * `Bearer file-mode` header, from a state the user believes is offline
   * (finding 16 of the 2026-08-05 security review).
   */
  const applyTransport = useCallback((entry: SourceEntry | undefined): string | null => {
    if (entry?.type === 'file') {
      const doc = docMap.current.get(entry.id);
      if (!doc) {
        setAuthBypass(false);
        setGraphTransport(null);
        return 'That file is no longer loaded in this session. Load it again to inspect it.';
      }
      setAuthBypass(true);
      setGraphTransport(createFileTransport(doc));
      return null;
    }
    // No source, or the live source: real network transport, real auth.
    setAuthBypass(false);
    setGraphTransport(null);
    return null;
  }, []);

  // ── public API ────────────────────────────────────────────────────────────

  const activateSource = useCallback((id: string) => {
    setSourceError(applyTransport(sources.find(s => s.id === id)));
    setActiveSourceId(id);
    // No queryClient.clear() — each source has its own cache slot (sourceId
    // is part of the query key). Switching sources reuses the cached data for
    // the target source when it is still fresh, and only fetches when stale.
  }, [sources, applyTransport]);

  const addFileSource = useCallback((text: string, fileName: string) => {
    const doc = parseFixture(text); // throws FixtureParseError on bad input
    const id = nextFileId();
    docMap.current.set(id, doc);
    const entry: SourceEntry = { id, type: 'file', label: fileName };
    setSourceError(null);
    setAuthBypass(true);
    setGraphTransport(createFileTransport(doc));
    setSources(prev => [...prev, entry]);
    setActiveSourceId(id);
    // No queryClient.clear() — the new file id has no existing cache entry.
  }, []);

  const addLiveSource = useCallback((label: string) => {
    setSources(prev => {
      if (prev.some(s => s.type === 'live')) {
        // Already registered — just update the label in case the account changed.
        return prev.map(s => s.type === 'live' ? { ...s, label } : s);
      }
      return [...prev, { id: LIVE_ID, type: 'live', label }];
    });
    setSourceError(null);
    setAuthBypass(false);
    setGraphTransport(null);
    setActiveSourceId(LIVE_ID);
    // Drop the live cache on (re-)login so a new session always gets fresh data
    // (guards against a different tenant signing in after sign-out).
    queryClient.removeQueries({ queryKey: ['securityProfiles', LIVE_ID] });
  }, [queryClient]);

  const removeSource = useCallback((id: string) => {
    docMap.current.delete(id);
    const next = sources.filter(s => s.id !== id);
    setSources(next);
    if (activeSourceId === id) {
      setSourceError(applyTransport(next[0]));
      setActiveSourceId(next[0]?.id ?? null);
    }
    // Drop only the removed source's cache slot.
    queryClient.removeQueries({ queryKey: ['securityProfiles', id] });
  }, [sources, activeSourceId, applyTransport, queryClient]);

  // ── backward-compat derivations ───────────────────────────────────────────

  const activeEntry = sources.find(s => s.id === activeSourceId);
  const mode: DataSourceMode = activeEntry?.type === 'file' ? 'file' : 'graph';
  const fileName = activeEntry?.type === 'file' ? activeEntry.label : null;

  const enterFileMode = addFileSource;
  const exitFileMode = useCallback(() => {
    if (activeSourceId) removeSource(activeSourceId);
  }, [activeSourceId, removeSource]);

  const value = useMemo<DataSourceContextValue>(
    () => ({
      sources, activeSourceId, sourceError,
      activateSource, addFileSource, addLiveSource, removeSource,
      mode, fileName, enterFileMode, exitFileMode,
    }),
    [sources, activeSourceId, sourceError, activateSource, addFileSource,
     addLiveSource, removeSource, mode, fileName, enterFileMode, exitFileMode],
  );

  return (
    <DataSourceContext.Provider value={value}>
      {children}
    </DataSourceContext.Provider>
  );
}
