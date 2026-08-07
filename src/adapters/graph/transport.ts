/**
 * Pluggable transport seam — Tier 3 (bottom).
 *
 * `client.ts` performs every Graph round-trip through `getGraphFetch()` rather
 * than calling the global `fetch` directly. By default this is the real network
 * transport. In file mode (see docs/architecture.md §6.4) a file transport is
 * installed that answers from a local JSON fixture instead of the network — so
 * the connector, loader, mapper, and audit interceptor all run unchanged.
 *
 * Status: the seam is implemented here. Wiring `client.ts` to call
 * `getGraphFetch()` is part of the implementation step, pending team review.
 *
 * See docs/architecture.md §6.4 for the full contract.
 */

/** A fetch-compatible function. The file transport returns synthetic Responses. */
export type GraphFetch = (url: string, init?: RequestInit) => Promise<Response>;

let activeTransport: GraphFetch | null = null;

/**
 * Install a transport. Pass `null` to restore the default HTTP transport
 * (used when leaving file mode).
 */
export function setGraphTransport(transport: GraphFetch | null): void {
  activeTransport = transport;
}

/**
 * The transport `client.ts` should use for every Graph round-trip.
 * Falls back to the global `fetch` when no transport is installed.
 */
export function getGraphFetch(): GraphFetch {
  return activeTransport ?? ((url, init) => fetch(url, init));
}
