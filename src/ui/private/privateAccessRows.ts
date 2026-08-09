/**
 * Private Access view-model builder — Tier 1 (UI), pure functions.
 *
 * Transforms the read-only Private Access domain (apps, App Proxy apps,
 * auth-strength) plus the shared Conditional Access detail into display rows.
 * No Graph calls, no side effects — mirrors the pattern in policyRows.ts.
 *
 * Conditional Access correlation: an app is "covered" by a CA
 * policy when the policy targets the app's `appId` directly, or targets `All`
 * apps (minus explicit exclusions). `applicationFilter` (custom-security-
 * attribute) targeting is surfaced as "filter-based" but not evaluated here.
 *
 * IMPORTANT — Private Access app authorization vs. GSA security profiles: a CA policy that carries a
 * `sessionControls.globalSecureAccessFilteringProfile` (or the legacy
 * `networkAccessSecurity`) session control attaches a GSA *security/filtering
 * profile* to internet traffic — it is an **Internet Access** policy, not a
 * Private Access **app authorization** policy. Such policies are **excluded**
 * from per-app CA coverage so the Private Access view only shows the policies
 * that actually govern access to the app (target resource + grant controls).
 */

import type { ConditionalAccessPolicy } from '../../model/definitions/ConditionalAccessPolicy.definition';
import type {
  PrivateAccessApp,
  AppProxyApp,
  AuthenticationStrengthPolicy,
  PrivateAccessKind,
  ApplicationSegment,
} from '../../model/definitions/PrivateAccessApp.definition';
import { privateAccessKind } from '../../model/definitions/PrivateAccessApp.definition';
import { debugGroup, debugTable, isDebugEnabled } from '../../app/debugLog';

// ---------------------------------------------------------------------------
// Built-in Conditional Access control labels
// ---------------------------------------------------------------------------

const BUILTIN_CONTROL_LABEL: Record<string, string> = {
  mfa: 'Require MFA',
  block: 'Block access',
  compliantDevice: 'Require compliant device',
  domainJoinedDevice: 'Require Entra hybrid joined device',
  approvedApplication: 'Require approved client app',
  compliantApplication: 'Require app protection policy',
  passwordChange: 'Require password change',
};

/** How a CA policy reaches an app. */
export type CoverageKind = 'direct' | 'all' | 'filter';

/** A CA policy that applies to a given app, with the relevant grant controls. */
export interface AppCaCoverage {
  id: string;
  name: string;
  state?: string;
  coverage: CoverageKind;
  /** Friendly grant-control labels (e.g. "Require MFA"). */
  controls: string[];
  /** Resolved authentication-strength name, when the policy references one. */
  authStrength?: string;
  /** Effective access: `block` when a Block grant control is present, else `grant`. */
  access: 'block' | 'grant';
}

/** A Private Access / Quick Access application display row. */
export interface PrivateAccessRow {
  id: string;
  appId: string;
  name: string;
  kind: PrivateAccessKind;
  hasCustomSecurityAttributes: boolean;
  coverages: AppCaCoverage[];
  /** Per-app network segments (destinations + ports). Empty when none captured. */
  segments: SegmentRow[];
  /** Private DNS posture for the app (ztspecs 25399/25400). */
  dns: DnsStatus;
  /** Who can reach the app (ztspecs 25480/25481). */
  assignment: AssignmentSummary;
  /** Lower-cased haystack for global search. */
  search: string;
}

/** An Application Proxy application display row. */
export interface AppProxyRow {
  id: string;
  appId: string;
  name: string;
  /** `aadPreAuthentication` | `passthru` | other. */
  externalAuthenticationType?: string;
  /** True when pre-auth (Entra) is enforced. */
  preAuthEnforced: boolean;
  externalUrl?: string;
  internalUrl?: string;
  coverages: AppCaCoverage[];
  search: string;
}

export interface PrivateAccessViewModel {
  apps: PrivateAccessRow[];
  appProxyApps: AppProxyRow[];
  /** Count of apps with at least one covering CA policy. */
  coveredAppCount: number;
}

// ---------------------------------------------------------------------------
// Application segments (per-app destinations) — ztspecs 25395
// ---------------------------------------------------------------------------

/**
 * Destination breadth of a segment, used purely for visual inspection (not a
 * pass/fail verdict). `wildcard` = a `dnsSuffix` segment; `broad` = a CIDR
 * shorter than /24 or an IP range spanning more than 256 addresses; otherwise
 * `granular`. See ztspecs 25395 "Segment Granularity Check".
 */
export type SegmentBreadth = 'granular' | 'broad' | 'wildcard';

/** A single application segment projected for display. */
export interface SegmentRow {
  id: string;
  destinationHost: string;
  destinationType: string;
  /** Friendly destination-type label (e.g. "IP range", "DNS suffix"). */
  destinationTypeLabel: string;
  /** Formatted port list (e.g. "445, 3389, 49152-65535"); "—" when none. */
  ports: string;
  protocol: string;
  action: string;
  breadth: SegmentBreadth;
  /** True for a DNS-suffix segment (Private DNS). */
  isDnsSuffix: boolean;
  /** True when port 53 is published (DNS over the tunnel). */
  hasPort53: boolean;
  /** True when the segment opens more than 10 ports (AD ephemeral ranges exempt). */
  widePorts: boolean;
}

/** Per-app Private DNS posture (ztspecs 25399 / 25400). */
export interface DnsStatus {
  /** `onPremisesPublishing.isDnsResolutionEnabled`. */
  resolutionEnabled: boolean;
  /** Number of `dnsSuffix` segments configured. */
  suffixCount: number;
  /** True when at least one segment publishes port 53. */
  hasPort53: boolean;
  /** True when DNS resolution is enabled OR a suffix/port-53 segment exists. */
  configured: boolean;
}

/** Who can reach an app (ztspecs 25480/25481). */
export interface AssignmentSummary {
  /** `appRoleAssignmentRequired` — when false, all users have implicit access. */
  required: boolean;
  /** Number of explicit user/group assignments. */
  count: number;
  /** Up-to-a-few assigned principals for display. */
  principals: { label: string; type: string }[];
  /** True when assignment is required but nothing is assigned (access gap). */
  hasGap: boolean;
}

const DESTINATION_TYPE_LABEL: Record<string, string> = {
  ip: 'IP address',
  ipRange: 'IP range',
  ipRangeCidr: 'CIDR',
  fqdn: 'FQDN',
  dnsSuffix: 'DNS suffix',
};

/** AD ephemeral RPC ranges that ztspecs 25395 exempts from "broad ports". */
const AD_EPHEMERAL_RANGES: ReadonlyArray<[number, number]> = [
  [49152, 65535],
  [1025, 5000],
];

function ipv4ToInt(ip: string): number | null {
  const parts = ip.trim().split('.');
  if (parts.length !== 4) return null;
  let n = 0;
  for (const part of parts) {
    const o = Number(part);
    if (!Number.isInteger(o) || o < 0 || o > 255) return null;
    n = n * 256 + o;
  }
  return n >>> 0;
}

/** Size (address count) of an `a..b` IP range, or null if unparseable. */
function ipRangeSize(host: string): number | null {
  const sep = host.includes('..') ? '..' : host.includes('-') ? '-' : null;
  if (!sep) return null;
  const [a, b] = host.split(sep);
  const lo = ipv4ToInt(a);
  const hi = ipv4ToInt(b);
  if (lo == null || hi == null || hi < lo) return null;
  return hi - lo + 1;
}

/** CIDR prefix length (e.g. 24 for `10.0.0.0/24`), or null. */
function cidrPrefix(host: string): number | null {
  const slash = host.lastIndexOf('/');
  if (slash < 0) return null;
  const p = Number(host.slice(slash + 1));
  return Number.isInteger(p) && p >= 0 && p <= 32 ? p : null;
}

function classifyBreadth(seg: ApplicationSegment): SegmentBreadth {
  const type = seg.destinationType ?? '';
  const host = seg.destinationHost ?? '';
  if (type === 'dnsSuffix') return 'wildcard';
  if (type === 'ipRangeCidr') {
    const prefix = cidrPrefix(host);
    return prefix != null && prefix < 24 ? 'broad' : 'granular';
  }
  if (type === 'ipRange') {
    const size = ipRangeSize(host);
    return size != null && size > 256 ? 'broad' : 'granular';
  }
  return 'granular';
}

/** Parse a `"start-end"` (or `"port"`) string into a numeric range. */
function parsePortRange(p: string): [number, number] | null {
  const m = p.split('-');
  const lo = Number(m[0]);
  const hi = m.length > 1 ? Number(m[1]) : lo;
  if (!Number.isInteger(lo) || !Number.isInteger(hi)) return null;
  return [lo, hi];
}

function isAdEphemeral(lo: number, hi: number): boolean {
  return AD_EPHEMERAL_RANGES.some(([a, b]) => lo === a && hi === b);
}

/** Collapse `"445-445"` → `"445"`, keep `"49152-65535"` as a range. */
function formatPorts(ports: string[] | undefined): string {
  if (!ports || ports.length === 0) return '—';
  return ports
    .map((p) => {
      const r = parsePortRange(p);
      if (!r) return p;
      return r[0] === r[1] ? String(r[0]) : `${r[0]}-${r[1]}`;
    })
    .join(', ');
}

function hasPort53(seg: ApplicationSegment): boolean {
  for (const p of seg.ports ?? []) {
    const r = parsePortRange(p);
    if (r && r[0] <= 53 && 53 <= r[1]) return true;
  }
  return false;
}

/** Total opened ports, treating AD ephemeral RPC ranges as a single "slot". */
function hasWidePorts(seg: ApplicationSegment): boolean {
  let count = 0;
  for (const p of seg.ports ?? []) {
    const r = parsePortRange(p);
    if (!r) continue;
    if (isAdEphemeral(r[0], r[1])) continue; // AD RPC range — exempt (ztspecs 25395)
    count += r[1] - r[0] + 1;
  }
  return count > 10;
}

function toSegmentRow(seg: ApplicationSegment): SegmentRow {
  const destinationType = seg.destinationType ?? 'unknown';
  return {
    id: seg.id ?? `${seg.destinationHost ?? ''}-${(seg.ports ?? []).join('_')}`,
    destinationHost: seg.destinationHost ?? '—',
    destinationType,
    destinationTypeLabel: DESTINATION_TYPE_LABEL[destinationType] ?? destinationType,
    ports: formatPorts(seg.ports),
    protocol: seg.protocol && seg.protocol !== '0' ? seg.protocol : '—',
    action: seg.action ?? 'tunnel',
    breadth: classifyBreadth(seg),
    isDnsSuffix: destinationType === 'dnsSuffix',
    hasPort53: hasPort53(seg),
    widePorts: hasWidePorts(seg),
  };
}

// ---------------------------------------------------------------------------
// Segment ↔ destination matching (What-If for Private Access — V1 item 4)
//
// The query side supports the patterns a network admin actually types, and a
// match is an OVERLAP between the set the query describes and the set the
// segment covers (so a CIDR query finds every segment inside or straddling it):
//
//   • IP            10.20.30.40
//   • CIDR          192.168.110.0/24
//   • IP range      192.168.110.10..192.168.110.20   (or a-b)
//   • IP wildcard   192.168.110.*  ·  192.168.*
//   • IP prefix     192.168.110            (treated as 192.168.110.0/24)
//   • FQDN          app.contoso.com
//   • FQDN glob     *.contoso.com  ·  www.contoso*  ·  *contoso*
//
// Inspection heuristic only — not a DNS resolver or live-traffic sim (spec §7).
// IP-family queries never match name segments and vice-versa.
// ---------------------------------------------------------------------------

/** Which segment dimension matched a What-If destination query. */
export type SegmentMatchReason = 'fqdn' | 'dns-suffix' | 'ip' | 'ip-range' | 'cidr';

/** An inclusive [low, high] IPv4 interval as uint32 numbers. */
type IpInterval = [number, number];

/**
 * Strip scheme + path from a query, leaving the bare host / ip / pattern.
 * Keeps a trailing `/<n>` when it is a CIDR prefix (1–2 digits) rather than a
 * URL path segment.
 */
function toHostQuery(query: string): string {
  let s = query.trim().toLowerCase().replace(/^[a-z][a-z0-9+.-]*:\/\//, '');
  const slash = s.indexOf('/');
  if (slash >= 0 && !/^\d{1,2}$/.test(s.slice(slash + 1))) {
    s = s.slice(0, slash);
  }
  return s;
}

function octetsToInt(octets: number[]): number {
  let n = 0;
  for (const o of octets) n = n * 256 + o;
  return n >>> 0;
}

/** [low, high] interval covered by a CIDR string (e.g. `10.0.0.0/8`). */
function cidrInterval(host: string): IpInterval | null {
  const slash = host.lastIndexOf('/');
  if (slash < 0) return null;
  const base = ipv4ToInt(host.slice(0, slash));
  const prefix = Number(host.slice(slash + 1));
  if (base == null || !Number.isInteger(prefix) || prefix < 0 || prefix > 32) return null;
  const mask = prefix === 0 ? 0 : (0xffffffff << (32 - prefix)) >>> 0;
  const lo = (base & mask) >>> 0;
  const hi = (lo | (~mask >>> 0)) >>> 0;
  return [lo, hi];
}

/** [low, high] interval for an `a..b` (or `a-b`) IPv4 range string. */
function rangeInterval(host: string): IpInterval | null {
  const sep = host.includes('..') ? '..' : host.includes('-') ? '-' : null;
  if (!sep) return null;
  const [a, b] = host.split(sep);
  const lo = ipv4ToInt(a);
  const hi = ipv4ToInt(b);
  if (lo == null || hi == null) return null;
  return lo <= hi ? [lo, hi] : [hi, lo];
}

/**
 * [low, high] interval for an IP wildcard / prefix (`192.168.110.*`,
 * `192.168.*`, or a bare prefix `192.168.110`). Missing or `*` octets widen to
 * 0–255.
 */
function wildcardInterval(host: string): IpInterval | null {
  const parts = host.split('.');
  if (parts.length === 0 || parts.length > 4) return null;
  const lo: number[] = [];
  const hi: number[] = [];
  for (let i = 0; i < 4; i++) {
    const p = i < parts.length ? parts[i] : '*';
    if (p === '*' || p === '') {
      lo.push(0);
      hi.push(255);
    } else {
      const n = Number(p);
      if (!Number.isInteger(n) || n < 0 || n > 255) return null;
      lo.push(n);
      hi.push(n);
    }
  }
  return [octetsToInt(lo), octetsToInt(hi)];
}

/** Parse a query into an IPv4 interval, or null when it is name-based. */
function queryIpInterval(q: string): IpInterval | null {
  // Letters → FQDN; no digits at all (e.g. bare `*`) → treat as FQDN glob.
  if (/[a-z]/i.test(q) || !/[0-9]/.test(q)) return null;
  if (q.includes('/')) return cidrInterval(q);
  if (q.includes('..') || q.includes('-')) {
    const r = rangeInterval(q);
    if (r) return r;
  }
  return wildcardInterval(q);
}

/** [low, high] interval a segment covers, or null when it is name-based. */
function segmentIpInterval(seg: SegmentRow): IpInterval | null {
  const dest = (seg.destinationHost ?? '').toLowerCase();
  switch (seg.destinationType) {
    case 'ip': {
      const d = ipv4ToInt(dest);
      return d == null ? null : [d, d];
    }
    case 'ipRange':
      return rangeInterval(dest);
    case 'ipRangeCidr':
      return cidrInterval(dest);
    default:
      return null;
  }
}

function intervalsOverlap(a: IpInterval, b: IpInterval): boolean {
  return a[0] <= b[1] && b[0] <= a[1];
}

/** DNS names cap at 253 chars; anything longer is not a real destination. */
const MAX_HOST_LEN = 253;

/**
 * Match a host glob (`*` = any run of characters) in linear time.
 *
 * Deliberately not a RegExp: compiling `*` to `.*` produces catastrophic
 * backtracking on inputs like `*a*a*a*a*b`, and this runs against every segment
 * on every keystroke. The two-pointer walk with a single backtrack point is
 * O(pattern x text) worst case and cannot blow up.
 *
 * Exported for the security regression suite (`tests/security/privateAccessRows.test.ts`).
 */
export function globMatches(pattern: string, text: string): boolean {
  const p = pattern.toLowerCase();
  const t = text.toLowerCase();
  let pi = 0;
  let ti = 0;
  let star = -1;
  let mark = 0;

  while (ti < t.length) {
    if (pi < p.length && (p[pi] === t[ti] || p[pi] === '?')) {
      pi++;
      ti++;
    } else if (pi < p.length && p[pi] === '*') {
      star = pi++;
      mark = ti;
    } else if (star >= 0) {
      pi = star + 1;
      ti = ++mark;
    } else {
      return false;
    }
  }
  while (pi < p.length && p[pi] === '*') pi++;
  return pi === p.length;
}

/**
 * True when an FQDN-family query overlaps a name segment. For a `dnsSuffix`
 * segment the covered set is `suffix` plus any `*.suffix`, tested in both
 * directions so both broad queries (`*.contoso.com`) and concrete queries
 * (`app.contoso.com`) match the suffix.
 */
function fqdnQueryMatches(q: string, segHost: string, isSuffix: boolean): boolean {
  if (q.length > MAX_HOST_LEN || segHost.length > MAX_HOST_LEN) return false;
  if (!isSuffix) return globMatches(q, segHost);

  // Query describes the suffix or a subdomain of it.
  if (globMatches(q, segHost) || globMatches(q, `host.${segHost}`)) return true;

  // Concrete query host that falls under the suffix.
  const concrete = q.replace(/\*/g, 'x');
  const seg = segHost.toLowerCase();
  const lower = concrete.toLowerCase();
  return lower === seg || lower.endsWith(`.${seg}`);
}

/**
 * Match a single application segment against a What-If destination query.
 * Returns the segment dimension that matched, or `null` when the segment does
 * not overlap the query.
 */
export function matchSegmentDestination(
  seg: SegmentRow,
  query: string,
): SegmentMatchReason | null {
  const q = toHostQuery(query);
  if (q === '') return null;

  const qInterval = queryIpInterval(q);
  if (qInterval) {
    const segInterval = segmentIpInterval(seg);
    if (!segInterval || !intervalsOverlap(qInterval, segInterval)) return null;
    return seg.destinationType === 'ipRangeCidr'
      ? 'cidr'
      : seg.destinationType === 'ipRange'
        ? 'ip-range'
        : 'ip';
  }

  // FQDN-family query.
  const dest = (seg.destinationHost ?? '').toLowerCase();
  if (seg.destinationType === 'fqdn') {
    return fqdnQueryMatches(q, dest, false) ? 'fqdn' : null;
  }
  if (seg.destinationType === 'dnsSuffix') {
    return fqdnQueryMatches(q, dest, true) ? 'dns-suffix' : null;
  }
  return null;
}

function buildDnsStatus(app: PrivateAccessApp, segments: SegmentRow[]): DnsStatus {
  const resolutionEnabled = app.isDnsResolutionEnabled === true;
  const suffixCount = segments.filter((s) => s.isDnsSuffix).length;
  const port53 = segments.some((s) => s.hasPort53);
  return {
    resolutionEnabled,
    suffixCount,
    hasPort53: port53,
    configured: resolutionEnabled || suffixCount > 0 || port53,
  };
}

const ASSIGNMENT_PRINCIPAL_LIMIT = 6;

function buildAssignmentSummary(app: PrivateAccessApp): AssignmentSummary {
  const assigned = app.appRoleAssignedTo ?? [];
  // Only flag as required when the SP explicitly says so.
  // When appRoleAssignmentRequired is undefined (SP data not yet fetched or
  // Directory.Read.All absent) we do NOT assume required — that would produce
  // false-positive red cards for every app where SP correlation is missing.
  const required = app.appRoleAssignmentRequired === true;
  const principals = assigned
    .slice(0, ASSIGNMENT_PRINCIPAL_LIMIT)
    .map((a) => ({
      label: a.principalDisplayName ?? a.principalId ?? 'Unknown',
      type: a.principalType ?? 'Unknown',
    }));
  return {
    required,
    count: assigned.length,
    principals,
    hasGap: required && assigned.length === 0,
  };
}

// ---------------------------------------------------------------------------
// Grant-control extraction
// ---------------------------------------------------------------------------

function controlLabels(ca: ConditionalAccessPolicy): string[] {
  const builtIn = ca.grantControls?.builtInControls ?? [];
  return builtIn.map((c) => BUILTIN_CONTROL_LABEL[c] ?? c);
}

/** Read `grantControls.authenticationStrength.id` without widening the type. */
function authStrengthId(ca: ConditionalAccessPolicy): string | undefined {
  const grant = ca.grantControls as { authenticationStrength?: { id?: string } } | undefined;
  return grant?.authenticationStrength?.id;
}

/**
 * True when a CA policy is a GSA **security-profile attachment** — i.e. it
 * carries a `globalSecureAccessFilteringProfile` (or the legacy
 * `networkAccessSecurity`) session control. Such a policy attaches a GSA
 * security/filtering profile to internet traffic (Internet Access / SWG
 * enforcement); it does **not** authorise access to a specific Private Access
 * enterprise app, so it is excluded from per-app CA coverage.
 */
function isSecurityProfileAttachment(ca: ConditionalAccessPolicy): boolean {
  const sc = ca.sessionControls;
  return Boolean(
    sc?.globalSecureAccessFilteringProfile?.isEnabled ||
      sc?.networkAccessSecurity?.isEnabled,
  );
}

// ---------------------------------------------------------------------------
// Coverage resolution
// ---------------------------------------------------------------------------

function coverageFor(
  appId: string,
  appLabel: string,
  caDetails: ConditionalAccessPolicy[],
  authStrengthById: Map<string, AuthenticationStrengthPolicy>,
): AppCaCoverage[] {
  const out: AppCaCoverage[] = [];
  const tracing = isDebugEnabled();
  const trace: Array<Record<string, unknown>> = [];

  for (const ca of caDetails) {
    const name = ca.displayName ?? ca.id;

    // A GSA security-profile attachment policy governs internet traffic, not
    // Private Access app access — exclude it from this app's coverage.
    if (isSecurityProfileAttachment(ca)) {
      if (tracing)
        trace.push({
          policy: name,
          decision: 'excluded',
          reason: 'GSA security-profile attachment (Internet Access, not PA app)',
        });
      continue;
    }

    const apps = ca.conditions?.applications;
    if (!apps) {
      if (tracing)
        trace.push({ policy: name, decision: 'skip', reason: 'no application condition' });
      continue;
    }

    const include = apps.includeApplications ?? [];
    const exclude = apps.excludeApplications ?? [];
    if (exclude.includes(appId)) {
      if (tracing)
        trace.push({ policy: name, decision: 'excluded', reason: 'app in excludeApplications' });
      continue;
    }

    let coverage: CoverageKind | null = null;
    if (include.includes(appId)) {
      coverage = 'direct';
    } else if (include.includes('All')) {
      coverage = 'all';
    } else if (
      (ca.conditions?.applications as { applicationFilter?: unknown } | undefined)
        ?.applicationFilter
    ) {
      // Filter-based (custom security attribute) targeting — surfaced, not evaluated.
      coverage = 'filter';
    }
    if (!coverage) {
      if (tracing) {
        const preview = include.slice(0, 4).join(', ') + (include.length > 4 ? '…' : '');
        trace.push({
          policy: name,
          decision: 'no-match',
          reason: `app not targeted (includeApplications=[${preview}])`,
        });
      }
      continue;
    }

    const strengthId = authStrengthId(ca);
    const controls = controlLabels(ca);
    out.push({
      id: ca.id,
      name,
      state: ca.state,
      coverage,
      controls,
      authStrength: strengthId ? authStrengthById.get(strengthId)?.displayName : undefined,
      access: controls.includes('Block access') ? 'block' : 'grant',
    });
    if (tracing)
      trace.push({
        policy: name,
        decision: `match: ${coverage}`,
        state: ca.state ?? '',
        controls: controls.join(', ') || '—',
      });
  }

  // Direct hits first, then All, then filter-based.
  const rank: Record<CoverageKind, number> = { direct: 0, all: 1, filter: 2 };
  out.sort((a, b) => rank[a.coverage] - rank[b.coverage]);

  debugGroup(
    `[GSA] CA matching · ${appLabel} (${appId}) → ${out.length} of ${caDetails.length} policies`,
    () => {
      debugTable(trace);
    },
  );

  return out;
}

// ---------------------------------------------------------------------------
// Public builder
// ---------------------------------------------------------------------------

export function buildPrivateAccessViewModel(
  apps: PrivateAccessApp[],
  appProxyApps: AppProxyApp[],
  authStrength: AuthenticationStrengthPolicy[],
  caDetails: ConditionalAccessPolicy[],
): PrivateAccessViewModel {
  const authStrengthById = new Map(authStrength.map((p) => [p.id, p]));

  const appRows: PrivateAccessRow[] = apps.map((app) => {
    const appId = app.appId ?? '';
    const name = app.displayName ?? (appId || app.id);
    const coverages = appId ? coverageFor(appId, name, caDetails, authStrengthById) : [];
    const csa =
      app.customSecurityAttributes != null &&
      Object.keys(app.customSecurityAttributes).length > 0;
    const segments = (app.applicationSegments ?? []).map(toSegmentRow);
    const dns = buildDnsStatus(app, segments);
    const assignment = buildAssignmentSummary(app);
    return {
      id: app.id,
      appId,
      name,
      kind: privateAccessKind(app),
      hasCustomSecurityAttributes: csa,
      coverages,
      segments,
      dns,
      assignment,
      search: [
        name,
        appId,
        app.id,
        ...(app.tags ?? []),
        ...segments.map((s) => s.destinationHost),
        ...assignment.principals.map((p) => p.label),
      ]
        .join(' ')
        .toLowerCase(),
    };
  });

  const proxyRows: AppProxyRow[] = appProxyApps.map((app) => {
    const appId = app.appId ?? '';
    const pub = app.onPremisesPublishing;
    const externalAuthenticationType = pub?.externalAuthenticationType;
    const name = app.displayName ?? app.id;
    const coverages = appId ? coverageFor(appId, name, caDetails, authStrengthById) : [];
    return {
      id: app.id,
      appId,
      name,
      externalAuthenticationType,
      preAuthEnforced: externalAuthenticationType === 'aadPreAuthentication',
      externalUrl: pub?.externalUrl,
      internalUrl: pub?.internalUrl,
      coverages,
      search: [name, appId, externalAuthenticationType, pub?.externalUrl]
        .filter(Boolean)
        .join(' ')
        .toLowerCase(),
    };
  });

  const coveredAppCount =
    appRows.filter((r) => r.coverages.length > 0).length +
    proxyRows.filter((r) => r.coverages.length > 0).length;

  return { apps: appRows, appProxyApps: proxyRows, coveredAppCount };
}

// ---------------------------------------------------------------------------
// Forwarding-profile ↔ app join
// ---------------------------------------------------------------------------

/**
 * Returns the subset of `apps` that are listed in a `private` forwarding
 * profile's `linkedApps[]`. Keyed by `appId` — O(n) lookup.
 *
 * For profiles that carry no `linkedApps` (GA profiles or live Graph where the
 * field is absent), returns all apps as a best-effort fallback so the UI
 * degrades gracefully rather than showing an empty list.
 *
 * See docs/Support-Multiple-Forwarding-Profiles.md §4.2 for the design
 * rationale (linkedApps-only, no segment duplication).
 */
export function appsInForwardingProfile(
  profile: { linkedApps?: { appId: string }[] },
  apps: PrivateAccessApp[],
): PrivateAccessApp[] {
  if (!profile.linkedApps || profile.linkedApps.length === 0) return apps;
  const ids = new Set(profile.linkedApps.map((a) => a.appId));
  return apps.filter((app) => app.appId != null && ids.has(app.appId));
}

/** Friendly label for a Private Access app kind. */
export function privateAccessKindLabel(kind: PrivateAccessKind): string {
  switch (kind) {
    case 'privateAccess':
      return 'Private Access';
    case 'quickAccess':
      return 'Quick Access';
    default:
      return 'Application';
  }
}
