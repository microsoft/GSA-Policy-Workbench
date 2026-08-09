/**
 * Gap-report interceptor seam.
 *
 * Today this is a no-op event sink. It exists *now* so that the Graph adapter
 * routes every call through a single chokepoint from day one — when the gap
 * report module lands, it just subscribes here and every call ever written is
 * captured for free. Retrofitting this later is expensive; wiring the seam now
 * is nearly free.
 *
 * See the project brief: "build the interceptor first, the table second."
 */

export interface GraphCallRecord {
  /** ISO timestamp. */
  ts: string;
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  /** Endpoint *template*, not the populated URL (no ids, no tenant data). */
  endpoint: string;
  scopes: string[];
  status: number;
  durationMs: number;
  apiVersion: 'beta' | 'v1.0';
  /** Logical feature that triggered the call: 'load-table', 'where-used', ... */
  feature: string;
  /** Set when the adapter had to route around a documented gap. */
  fallback?: { reason: string; alternativePath: string };
  notes?: string;
}

type Listener = (record: GraphCallRecord) => void;

const listeners = new Set<Listener>();

/** Subscribe to Graph call records. Returns an unsubscribe function. */
export function onGraphCall(listener: Listener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** Emit a Graph call record to all subscribers. Safe no-op when none. */
export function recordGraphCall(record: GraphCallRecord): void {
  for (const listener of listeners) {
    try {
      listener(record);
    } catch {
      // A misbehaving listener must never break a Graph call.
    }
  }
}
