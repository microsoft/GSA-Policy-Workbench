/**
 * What-If for Private Access — Tier 1 (UI), pure logic (V1 item 4).
 *
 * Given a destination (FQDN / IP / IP range / CIDR / DNS suffix), report which
 * Private Access / Quick Access **application segments** cover it, and the
 * Conditional Access outcome of the covering app. This is the Private Access
 * counterpart to the Internet Access effective-policy resolver
 * (`effectivePolicy.ts`) and follows the same rule: segment match only, **not**
 * a live-traffic simulation.
 *
 * Operates on the already-built Private Access view-model rows (segments + CA
 * coverage) — no Graph calls, no side effects.
 */

import {
  matchSegmentDestination,
  type PrivateAccessRow,
  type SegmentMatchReason,
  type SegmentRow,
} from '../private/privateAccessRows';
import type { ForwardingProfile } from '../../model/definitions/ForwardingProfile.definition';
import { evaluateAcquisition, type AcquisitionEvaluation } from './effectivePolicy';

/** A segment within an app that matched the destination. */
export interface SegmentHit {
  segment: SegmentRow;
  reason: SegmentMatchReason;
}

/** An app whose segments cover the destination, with its CA outcome. */
export interface PrivateAccessAppMatch {
  app: PrivateAccessRow;
  hits: SegmentHit[];
  /** `block` when a covering CA policy blocks, `grant` when one grants, else `none`. */
  access: 'block' | 'grant' | 'none';
  /**
   * Set on a Quick Access match when an Enterprise App ALSO covers this same
   * destination. Per GSA's documented overlap tie-break: "the segment defined on the Enterprise App will be given priority
   * by the GSA service. No traffic ... will be processed by Quick Access."
   * The Quick Access match is still returned (never hidden — this app's
   * "always explain, never silently drop" convention) but callers should
   * render it as not actually applied, naming the superseding app(s).
   */
  supersededBy?: string[];
}

export interface PrivateAccessWhatIfResult {
  destination?: string;
  hasDestination: boolean;
  /** Apps with at least one matching segment, most-specific first. */
  matches: PrivateAccessAppMatch[];
  /**
   * Client-side traffic-forwarding acquisition verdict for the destination
   * against the Private Access forwarding profile (stage 1). `undefined` when no forwarding-profile data
   * was supplied to `resolvePrivateAccessDestination`.
   */
  acquisition?: AcquisitionEvaluation;
}

/** Specificity rank for ordering hits (lower = more specific). */
const REASON_RANK: Record<SegmentMatchReason, number> = {
  ip: 0,
  fqdn: 1,
  cidr: 2,
  'ip-range': 3,
  'dns-suffix': 4,
};

function appAccess(app: PrivateAccessRow): 'block' | 'grant' | 'none' {
  if (app.coverages.some((c) => c.access === 'block')) return 'block';
  if (app.coverages.some((c) => c.access === 'grant')) return 'grant';
  return 'none';
}

/**
 * Resolve which Private Access apps cover a destination. Empty result when no
 * destination is given (the panel then shows nothing for PA).
 */
export function resolvePrivateAccessDestination(
  rows: PrivateAccessRow[],
  destination: string | undefined,
  forwardingProfiles: ForwardingProfile[] = [],
): PrivateAccessWhatIfResult {
  const q = destination?.trim() || undefined;
  if (!q) return { hasDestination: false, matches: [] };

  const matches: PrivateAccessAppMatch[] = [];
  for (const app of rows) {
    const hits: SegmentHit[] = [];
    for (const segment of app.segments) {
      const reason = matchSegmentDestination(segment, q);
      if (reason) hits.push({ segment, reason });
    }
    if (hits.length > 0) {
      hits.sort((a, b) => REASON_RANK[a.reason] - REASON_RANK[b.reason]);
      matches.push({ app, hits, access: appAccess(app) });
    }
  }

  // Most-specific match first (apps with a more specific hit rank higher).
  matches.sort((a, b) => REASON_RANK[a.hits[0].reason] - REASON_RANK[b.hits[0].reason]);

  // Enterprise App > Quick Access overlap tie-break — annotate (never remove) any Quick Access match when an Enterprise
  // App also covers this same destination, since GSA never routes that
  // traffic through Quick Access.
  const enterpriseAppNames = matches
    .filter((m) => m.app.kind !== 'quickAccess')
    .map((m) => m.app.name);
  if (enterpriseAppNames.length > 0) {
    for (const m of matches) {
      if (m.app.kind === 'quickAccess') m.supersededBy = enterpriseAppNames;
    }
  }

  return {
    destination: q,
    hasDestination: true,
    matches,
    acquisition: evaluateAcquisition(forwardingProfiles, 'private', q),
  };
}
