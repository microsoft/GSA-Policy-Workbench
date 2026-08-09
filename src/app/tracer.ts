/**
 * General-purpose structured trace facility.
 *
 * Modules write TraceRecords to named channels; registered sinks (console,
 * in-memory buffer) consume them independently. Channels: whatif, ca, loader,
 * graph. The graph channel is reserved for the Graph call interceptor.
 *
 * The console sink mirrors the existing debugLog.ts toggle behaviour (off by
 * default in prod, always on in dev, persisted in localStorage).
 */

export type TraceChannel = 'whatif' | 'ca' | 'loader' | 'graph';

export type TraceDecision =
  | 'pass'      // checked, not eliminated
  | 'skip'      // checked and eliminated / not applicable
  | 'match'     // a rule / destination matched
  | 'no-match'  // checked but nothing matched
  | 'winner'    // the winning step
  | 'escape'    // traffic escaped GSA
  | 'preempted' // claimed by a higher-precedence workload
  | 'info';     // neutral / informational

export interface TraceRecord {
  channel: TraceChannel;
  /** Coarse grouping label for the UI (e.g. "acquisition", "profile", "rule"). */
  stage: string;
  /** Short display label for this step. */
  label: string;
  decision: TraceDecision;
  /** Human-readable explanation of the decision. */
  reason: string;
  detail?: Record<string, unknown>;
  timestamp: number;
}

// ── Sink interface ──────────────────────────────────────────────────────────

interface Sink {
  write(record: TraceRecord): void;
}

// ── Console sink ────────────────────────────────────────────────────────────

const DECISION_ICON: Record<TraceDecision, string> = {
  pass:      '✓',
  skip:      '✗',
  match:     '●',
  'no-match': '○',
  winner:    '★',
  escape:    '⚠',
  preempted: '↗',
  info:      'ℹ',
};

const STAGE_SEP = '/';

class ConsoleSink implements Sink {
  private enabled: boolean;

  // Console tracing echoes tenant detail, so it never survives a reload and is
  // on by default only in a dev build. Re-enable per session via __gsaTrace.on().
  constructor() {
    this.enabled = import.meta.env.DEV;
  }

  setEnabled(v: boolean): void {
    this.enabled = v;
  }

  isEnabled(): boolean { return this.enabled; }

  write(record: TraceRecord): void {
    if (!this.enabled) return;
    const icon = DECISION_ICON[record.decision];
    console.log(
      `%c[${record.channel}${STAGE_SEP}${record.stage}] ${icon} ${record.label}`,
      'color:#888',
      '—', record.reason,
      record.detail ? record.detail : '',
    );
  }
}

// ── Buffer sink ─────────────────────────────────────────────────────────────

class BufferSink implements Sink {
  private readonly buffers = new Map<TraceChannel, TraceRecord[]>();
  private readonly max: number;

  constructor(max = 1000) { this.max = max; }

  write(record: TraceRecord): void {
    if (!this.buffers.has(record.channel)) this.buffers.set(record.channel, []);
    const buf = this.buffers.get(record.channel)!;
    buf.push(record);
    if (buf.length > this.max) buf.shift();
  }

  read(channel: TraceChannel): readonly TraceRecord[] {
    return this.buffers.get(channel) ?? [];
  }

  clear(channel?: TraceChannel): void {
    if (channel) this.buffers.delete(channel);
    else this.buffers.clear();
  }
}

// ── Tracer ──────────────────────────────────────────────────────────────────

class Tracer {
  private readonly consoleSink = new ConsoleSink();
  private readonly bufferSink = new BufferSink();
  private readonly sinks: Sink[];

  constructor() {
    this.sinks = [this.consoleSink, this.bufferSink];
  }

  /** Enable or disable the console sink (persisted across reloads). */
  setConsole(enabled: boolean): void { this.consoleSink.setEnabled(enabled); }

  isConsoleEnabled(): boolean { return this.consoleSink.isEnabled(); }

  /** Write one or more records to all sinks. */
  trace(records: TraceRecord | TraceRecord[]): void {
    const list = Array.isArray(records) ? records : [records];
    for (const r of list) {
      for (const sink of this.sinks) sink.write(r);
    }
  }

  /** Read the current buffer for a channel. */
  read(channel: TraceChannel): readonly TraceRecord[] {
    return this.bufferSink.read(channel);
  }

  /** Clear the buffer for a channel (or all channels). */
  clear(channel?: TraceChannel): void {
    this.bufferSink.clear(channel);
  }
}

export const tracer = new Tracer();

interface GsaTraceHandle {
  readonly enabled: boolean;
  on(): void;
  off(): void;
  read(channel: TraceChannel): readonly TraceRecord[];
  clear(channel?: TraceChannel): void;
}

/** Install the `window.__gsaTrace` handle. Idempotent — call once at bootstrap. */
export function installTracer(): void {
  if (typeof window === 'undefined') return;
  const w = window as unknown as { __gsaTrace?: GsaTraceHandle };
  if (w.__gsaTrace) return;
  w.__gsaTrace = {
    get enabled() { return tracer.isConsoleEnabled(); },
    on()  { tracer.setConsole(true);  console.log('[GSA trace] console output enabled'); },
    off() { tracer.setConsole(false); console.log('[GSA trace] console output disabled'); },
    read: (ch: TraceChannel) => tracer.read(ch),
    clear: (ch?: TraceChannel) => tracer.clear(ch),
  };
}
