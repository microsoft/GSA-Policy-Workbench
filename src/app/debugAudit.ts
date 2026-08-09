/**
 * Debug / call-log inspection — V1 item 5.
 *
 * Makes the Graph adapter's interceptor call log visible for **live
 * troubleshooting**. The interceptor (`adapters/graph/interceptor.ts`) already
 * records every Graph call through a single seam; this module subscribes once,
 * keeps a bounded in-memory buffer, and exposes it through the browser's F12
 * devtools: a `window.__gsaAudit` handle plus a compact `console.debug` line as
 * each call happens. This is what helps most when debugging a live tenant
 * mid-session.
 *
 * Privacy: every record's `endpoint` is a **template** (`{id}` placeholders, no
 * tenant/user/policy ids or names), so the log is
 * safe to read aloud, paste, or attach.
 *
 * This is a read-only observer — it never affects Graph calls.
 */

import { onGraphCall, type GraphCallRecord } from '../adapters/graph/interceptor';

const MAX_RECORDS = 500;
const records: GraphCallRecord[] = [];
let installed = false;

interface GsaAuditHandle {
  readonly records: GraphCallRecord[];
  /** Pretty-print the log as a console table. Returns the record count. */
  table(): number;
  /** The log serialised as pretty JSON. */
  json(): string;
  /** Clear the buffer. */
  clear(): number;
}

function compactLine(r: GraphCallRecord): string {
  const fb = r.fallback ? ` ⤳ fallback: ${r.fallback.reason}` : '';
  return `[GSA] ${r.method} ${r.endpoint} → ${r.status} (${r.durationMs}ms · ${r.feature})${fb}`;
}

/** A snapshot copy of the current call log. */
export function getAuditLog(): GraphCallRecord[] {
  return [...records];
}

/**
 * Subscribe to the interceptor and install the `window.__gsaAudit` handle.
 * Idempotent — safe to call once at app bootstrap.
 */
export function installDebugAudit(): void {
  if (installed) return;
  installed = true;

  onGraphCall((record) => {
    records.push(record);
    if (records.length > MAX_RECORDS) records.shift();
    if (import.meta.env.DEV) {
      const style = record.status >= 400 ? 'color:#b91c1c' : 'color:#6b7280';
      console.debug(`%c${compactLine(record)}`, style);
    }
  });

  const handle: GsaAuditHandle = {
    get records() {
      return [...records];
    },
    table() {
      console.table(
        records.map((r) => ({
          method: r.method,
          endpoint: r.endpoint,
          status: r.status,
          ms: r.durationMs,
          feature: r.feature,
          api: r.apiVersion,
          fallback: r.fallback?.reason ?? '',
        })),
      );
      return records.length;
    },
    json() {
      return JSON.stringify(records, null, 2);
    },
    clear() {
      const n = records.length;
      records.length = 0;
      return n;
    },
  };

  (window as unknown as { __gsaAudit?: GsaAuditHandle }).__gsaAudit = handle;

  if (import.meta.env.DEV) {
    console.info(
      '%cGSA audit log ready — inspect Graph calls with window.__gsaAudit ' +
        '(.table(), .records, .json(), .clear())',
      'color:#0067b8;font-weight:bold',
    );
  }
}
