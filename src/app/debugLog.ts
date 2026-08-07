/**
 * Debug logging for policy evaluation — CA correlation + What-If.
 *
 * Separate from the Graph call-log (`debugAudit.ts` / `window.__gsaAudit`):
 * this traces the SPA's **in-memory reasoning** — which Conditional Access
 * policies matched a Private Access app and why, and how the What-If resolver
 * reached its outcome. Intended for diagnosing complex tenants in the F12
 * console (e.g. a CA policy that should cover an app but doesn't show up).
 *
 * Off by default to keep the console quiet. Turn it on for a session via the
 * `window.__gsaDebug` handle (`.on()` / `.off()`), which persists in
 * `localStorage` so it survives reloads. Always on in a dev build.
 *
 * Read-only: the logger never changes evaluation output — it only observes.
 */

const STORAGE_KEY = 'gsaDebug';

// Legacy: earlier builds persisted this flag, which meant a support session
// could leave tenant detail streaming to the console indefinitely. Clear it.
try {
  localStorage.removeItem(STORAGE_KEY);
} catch {
  /* ignore */
}

// On by default only in a dev build; an opt-in never survives a reload.
let enabled = import.meta.env.DEV;

/** Whether policy-evaluation debug logging is currently active. */
export function isDebugEnabled(): boolean {
  return enabled;
}

interface GsaDebugHandle {
  /** Current state. */
  readonly enabled: boolean;
  /** Enable logging (persists across reloads). */
  on(): void;
  /** Disable logging. */
  off(): void;
}

/** Log a collapsed console group, building its body via `fn`. No-op when off. */
export function debugGroup(label: string, fn: () => void): void {
  if (!enabled) return;
  console.groupCollapsed(label);
  try {
    fn();
  } finally {
    console.groupEnd();
  }
}

/** Log a line inside a debug group (or standalone). No-op when off. */
export function debugLog(...args: unknown[]): void {
  if (!enabled) return;
  console.log(...args);
}

/** Log a table inside a debug group. No-op when off. */
export function debugTable(rows: unknown[]): void {
  if (!enabled || rows.length === 0) return;
  console.table(rows);
}

/**
 * Install the `window.__gsaDebug` handle. Idempotent; call once at bootstrap.
 */
export function installDebugLog(): void {
  const handle: GsaDebugHandle = {
    get enabled() {
      return enabled;
    },
    on() {
      enabled = true;
      console.info(
        '[GSA] policy-evaluation debug logging ON for this session — CA matching + What-If will trace to the console.',
      );
    },
    off() {
      enabled = false;
      console.info('[GSA] policy-evaluation debug logging OFF.');
    },
  };

  (window as unknown as { __gsaDebug?: GsaDebugHandle }).__gsaDebug = handle;

  if (enabled) {
    console.info(
      '[GSA] policy-evaluation debug logging is ON. Use window.__gsaDebug.off() to silence, .on() to re-enable.',
    );
  }
}
