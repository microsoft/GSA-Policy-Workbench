/**
 * Data-source context — Tier 1 (file-based data source, see architecture.md §6.4).
 *
 * Holds the list of named data sources (one live Graph connection + any number
 * of fixture files) and which one is currently active. The active source drives
 * the pluggable transport seam in adapters/graph/transport.ts.
 *
 * Split from the provider component (DataSourceProvider.tsx) so this file
 * exports only non-components, satisfying `react-refresh/only-export-components`.
 */

import { createContext, useContext } from 'react';

export type DataSourceMode = 'graph' | 'file';

/** 'live' = authenticated Microsoft Graph; 'file' = local fixture. */
export type SourceType = 'live' | 'file';

export interface SourceEntry {
  id: string;
  type: SourceType;
  /** Tenant username for live sources; file name for file sources. */
  label: string;
}

export interface DataSourceContextValue {
  // ── Multi-source list ────────────────────────────────────────────────────
  sources: SourceEntry[];
  activeSourceId: string | null;
  /**
   * Set when the active source could not be activated — for example a file
   * source whose parsed document is no longer held in this session. The
   * transport is left unset and the auth bypass off when this is non-null.
   */
  sourceError: string | null;
  /** Switch the active transport to an already-registered source. */
  activateSource: (id: string) => void;
  /** Remove a source. If it was active, the next remaining source is activated. */
  removeSource: (id: string) => void;
  /** Parse a fixture file and register it as a new file source. Throws FixtureParseError. */
  addFileSource: (text: string, fileName: string) => void;
  /** Register (or re-register) the live Graph source. At most one live entry exists. */
  addLiveSource: (label: string) => void;

  // ── Backward-compat (derived from the active source) ────────────────────
  /** 'graph' when active source is live; 'file' when active source is a fixture. */
  mode: DataSourceMode;
  /** File name of the active fixture, or null when in live mode. */
  fileName: string | null;
  /** Alias for addFileSource — used by SignInGate and legacy callers. */
  enterFileMode: (text: string, fileName: string) => void;
  /** Remove the active file source (switches to live if available, else empty state). */
  exitFileMode: () => void;
}

export const DataSourceContext = createContext<DataSourceContextValue | undefined>(
  undefined,
);

export function useDataSource(): DataSourceContextValue {
  const ctx = useContext(DataSourceContext);
  if (!ctx) {
    throw new Error('useDataSource must be used within a DataSourceProvider');
  }
  return ctx;
}
