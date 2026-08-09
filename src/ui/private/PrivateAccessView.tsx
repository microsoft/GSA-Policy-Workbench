/**
 * Private Access view — Tier 1 (UI), read-only (V1 domain, spec §6).
 *
 * Renders the Private Access / Quick Access applications and Application Proxy
 * applications captured by the exporter as collapsible group cards — one per
 * application — each with an inner table of the Conditional Access policies that
 * cover it. The visual language deliberately mirrors the Internet Access policy
 * table (gradient group headers, zebra rule rows, the same block = danger
 * badge) so the two domains read consistently. Display only — no Graph calls.
 *
 * Data is sourced from the already-loaded ProfileTreeResult.privateAccess plus
 * the shared CA detail. In V0 this is populated only in file mode (the SPA does
 * not request the extra scopes the live endpoints need — spec §6.5), so the
 * view shows an explanatory empty state when there is nothing to display.
 */

import { useMemo, useState, type ReactElement } from 'react';
import {
  Badge,
  Button,
  Input,
  Text,
  ToggleButton,
  Tooltip,
  makeStyles,
  mergeClasses,
  tokens,
  type BadgeProps,
} from '@fluentui/react-components';
import {
  ChevronDown16Regular,
  ChevronRight16Regular,
  Search16Regular,
  LockClosed16Regular,
  LockClosed20Regular,
  AppGeneric20Regular,
  Globe20Regular,
  Tag16Regular,
  Globe16Regular,
  Filter16Regular,
  CheckmarkCircle16Filled,
  Warning16Filled,
  Open16Regular,
  Server16Regular,
  People16Regular,
  PersonProhibited16Regular,
  ShieldTask16Regular,
} from '@fluentui/react-icons';
import type { ConditionalAccessPolicy } from '../../model/definitions/ConditionalAccessPolicy.definition';
import type { PrivateAccessDomain } from '../../adapters/graph/loader';
import type { ForwardingProfile } from '../../model/definitions/ForwardingProfile.definition';
import {
  buildPrivateAccessViewModel,
  privateAccessKindLabel,
  type AppCaCoverage,
  type AppProxyRow,
  type PrivateAccessRow,
  type SegmentRow,
} from './privateAccessRows';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  search: { width: '340px' },
  count: { color: tokens.colorNeutralForeground3 },
  grow: { flex: 1 },
  scroll: {
    paddingRight: tokens.spacingHorizontalXS,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
  },

  // --- section ---------------------------------------------------------------
  section: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM },
  sectionHead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    color: tokens.colorNeutralForeground2,
  },
  sectionTitle: { fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightBold },
  /** Wraps the title row + meta row for profile-grouped sections. */
  profileHeadRow: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
  },
  /** Assignment + edge-fallback metadata line under a profile section header. */
  profileMeta: {
    display: 'flex',
    gap: tokens.spacingHorizontalL,
    paddingLeft: '28px', // visually aligns with text after the 20px icon
    flexWrap: 'wrap',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  fallbackChunk: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
  },

  // --- app group card (mirrors the IA profile card) --------------------------
  group: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusXLarge,
    overflow: 'hidden',
    boxShadow: tokens.shadow4,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  header: {
    display: 'grid',
    // Fixed swim-lane template — mirrors the IA profile card.
    // Lane 1 (identity) is a FIXED width holding TWO lines: the app name, and a
    // subtext row of app-specific badges (kind/segment/DNS/assignment/CSA) right
    // beneath it — the same "name, then a subtext line of secondary facts"
    // pattern IA uses for its "N policies · N rules" line. Folding the badges
    // into the identity lane (instead of giving them their own top-level grid
    // column) both pulls them close to the name AND frees up the lane sequence
    // so it mirrors IA's shape: identity | verdict | flexible trailing content.
    // The two fixed numbers below (516 and 200) are chosen so this header's
    // cumulative width up to the verdict lane (516 + gap = 528px) exactly
    // matches IA's (64 gutter + gap + 440 identity + gap = 528px), and the
    // verdict lane itself is the same 200px in both — so "blocked/allowed" (IA)
    // and "block/grant" (PA) land at the identical x, at any window width.
    gridTemplateColumns: '516px 200px minmax(0, 1fr) 176px',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    cursor: 'pointer',
    borderLeft: `4px solid ${tokens.colorBrandStroke1}`,
    background: `linear-gradient(90deg, ${tokens.colorBrandBackground2} 0%, ${tokens.colorNeutralBackground1} 70%)`,
    ':hover': {
      background: `linear-gradient(90deg, ${tokens.colorBrandBackground2Hover} 0%, ${tokens.colorNeutralBackground1Hover} 70%)`,
    },
  },
  headerWarn: {
    borderLeftColor: tokens.colorPaletteDarkOrangeBorderActive,
    background: `linear-gradient(90deg, ${tokens.colorPaletteDarkOrangeBackground2} 0%, ${tokens.colorNeutralBackground1} 70%)`,
    ':hover': {
      background: `linear-gradient(90deg, ${tokens.colorPaletteDarkOrangeBackground2} 0%, ${tokens.colorNeutralBackground1Hover} 70%)`,
    },
  },
  headerLead: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    minWidth: 0,
  },
  headerLeadTop: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  chevron: { display: 'inline-flex', color: tokens.colorNeutralForeground2 },
  appIcon: { display: 'inline-flex', color: tokens.colorBrandForeground1, flexShrink: 0 },
  appName: {
    fontWeight: tokens.fontWeightBold,
    fontSize: tokens.fontSizeBase400,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },

  // Fixed-width header slots so every badge group lines up in the same column
  // across cards — the collapsed list reads like a fixed-column table (item 6).
  cell: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
    overflow: 'hidden',
  },
  cellKind: { width: '120px' },
  cellSeg: { width: '108px' },
  cellDns: { width: '132px' },
  cellAssign: { width: '136px' },
  cellCsa: { width: '64px' },
  // Subtext row (line 2 of the identity lane) — the app-specific badge cluster,
  // directly beneath the name, same idea as IA's "N policies · N rules" subtext.
  headerMeta: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
    paddingLeft: '24px',
  },
  // Lane 2 — the block/grant verdict, fixed-width (200px, matching IA's action
  // lane) so it lands at the same x on every card, right after identity.
  headerVerdict: {
    display: 'flex',
    alignItems: 'center',
    overflow: 'hidden',
  },
  // Lane 4 — the low-emphasis CA-policy count. This is the one flexible track
  // (see the `.header` gridTemplateColumns comment) — it absorbs the leftover
  // width so lane 5 (appId) stays flush against the right edge, mirroring how
  // IA's CA-chips lane trails after its action lane.
  headerCaCount: {
    display: 'flex',
    alignItems: 'center',
    minWidth: 0,
    overflow: 'hidden',
  },
  headerRight: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    justifyContent: 'flex-end',
    minWidth: 0,
  },
  appId: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    maxWidth: '100%',
  },
  summaryRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
  },

  // --- inner CA coverage grid + application-segments grid --------------------
  // Both are rendered as a shared CSS Grid (one template reused by the title
  // row and every data row), not independent <table> elements — the same
  // anti-pattern fix documented for IA's expanded rule grid.
  // `minmax(0, …fr)` lanes are safe here because
  // every app's sub-table reuses the SAME class inside a container of the
  // SAME width, so a lane's resolved pixel width never depends on that
  // particular app's own content — Destination/Ports/Controls text length no
  // longer drags the Type/Protocol/Scope/Coverage/Access/State lanes sideways.
  gridWrap: {
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  covGridRow: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1.3fr) 150px 100px minmax(0, 1.5fr) 120px',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
  },
  segGridRow: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1.4fr) 170px minmax(0, 1.2fr) 90px 150px',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
  },
  gridHeaderRow: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    color: tokens.colorNeutralForeground3,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.03em',
  },
  gridDataRow: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  gridRowAlt: { backgroundColor: tokens.colorNeutralBackground2 },
  gridCell: { minWidth: 0, overflow: 'hidden' },
  gridCellWrap: { minWidth: 0, whiteSpace: 'normal', wordBreak: 'break-word' },
  muted: { color: tokens.colorNeutralForeground3 },
  cellBadge: { whiteSpace: 'nowrap' },
  controls: { color: tokens.colorNeutralForeground2 },
  subCaption: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.03em',
  },
  subCaptionToggle: {
    cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground2Hover },
  },
  mono: { fontFamily: tokens.fontFamilyMonospace },
  cellBadgeGap: { display: 'inline-flex', gap: tokens.spacingHorizontalXS, alignItems: 'center', flexWrap: 'wrap' },
  emptyCoverage: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },  noMatches: {
    padding: tokens.spacingVerticalXXL,
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  empty: {
    display: 'grid',
    placeItems: 'center',
    height: '100%',
    textAlign: 'center',
    gap: tokens.spacingVerticalS,
    color: tokens.colorNeutralForeground3,
  },
});

interface PrivateAccessViewProps {
  privateAccess: PrivateAccessDomain;
  caDetails: ConditionalAccessPolicy[];
  forwardingProfiles?: ForwardingProfile[];
}

// ---------------------------------------------------------------------------
// Badge helpers (kept consistent with the IA table)
// ---------------------------------------------------------------------------

/** Colour for the Access badge — block mirrors the IA "block" action (danger). */
function accessColor(access: 'block' | 'grant'): BadgeProps['color'] {
  return access === 'block' ? 'danger' : 'success';
}

/** Icon + colour for a coverage kind, shown in the Coverage column. */
const COVERAGE_META: Record<
  AppCaCoverage['coverage'],
  { label: string; icon: typeof Tag16Regular; color: BadgeProps['color'] }
> = {
  direct: { label: 'Targeted', icon: Tag16Regular, color: 'brand' },
  all: { label: 'All apps', icon: Globe16Regular, color: 'informative' },
  filter: { label: 'Filter-based', icon: Filter16Regular, color: 'warning' },
};

function stateBadge(state: string | undefined): { label: string; color: BadgeProps['color'] } | null {
  if (!state || state === 'enabled') return null;
  if (state === 'enabledForReportingButNotEnforced') return { label: 'Report-only', color: 'warning' };
  if (state === 'disabled') return { label: 'Disabled', color: 'subtle' };
  return { label: state, color: 'subtle' };
}

// ---------------------------------------------------------------------------
// Inner CA coverage table
// ---------------------------------------------------------------------------

function CoverageTable({ coverages }: { coverages: AppCaCoverage[] }) {
  const styles = useStyles();

  if (coverages.length === 0) {
    return (
      <div className={styles.emptyCoverage}>
        No covering Conditional Access policy found for this application.
      </div>
    );
  }

  return (
    <div className={styles.gridWrap}>
      <div className={mergeClasses(styles.covGridRow, styles.gridHeaderRow)}>
        <span>Conditional Access</span>
        <span>Coverage</span>
        <span>Access</span>
        <span>Controls</span>
        <span>State</span>
      </div>
      {coverages.map((c, i) => {
        const cov = COVERAGE_META[c.coverage];
        const CovIcon = cov.icon;
        const st = stateBadge(c.state);
        const controls = [...c.controls, c.authStrength].filter(Boolean).join(' · ');
        return (
          <div
            key={c.id}
            className={mergeClasses(
              styles.covGridRow,
              styles.gridDataRow,
              i % 2 === 1 && styles.gridRowAlt,
            )}
          >
            <span className={styles.gridCell}>{c.name}</span>
            <span className={styles.gridCell}>
              <Badge
                appearance="tint"
                color={cov.color}
                icon={<CovIcon />}
                className={styles.cellBadge}
              >
                {cov.label}
              </Badge>
            </span>
            <span className={styles.gridCell}>
              <Badge appearance="filled" color={accessColor(c.access)}>
                {c.access}
              </Badge>
            </span>
            <span className={mergeClasses(styles.gridCell, styles.gridCellWrap, styles.controls)}>
              {controls || '—'}
            </span>
            <span className={styles.gridCell}>
              {st ? (
                <Badge appearance="outline" color={st.color}>
                  {st.label}
                </Badge>
              ) : (
                <span className={styles.muted}>enabled</span>
              )}
            </span>
          </div>
        );
      })}
    </div>
  );
}

/** Sort blocked policies first, then grants — within each group, original order is kept. */
function sortCoveragesByAccess(coverages: AppCaCoverage[]): AppCaCoverage[] {
  const rank = (c: AppCaCoverage) => (c.access === 'block' ? 0 : 1);
  return [...coverages].sort((a, b) => rank(a) - rank(b));
}

/**
 * Collapsible wrapper around CoverageTable — collapsed by default (mirrors the
 * app-card default) with a short "N conditional access policies" header, so
 * expanding an app doesn't immediately dump a full CA table underneath the
 * application-segments table.
 */
function CoverageSection({ coverages }: { coverages: AppCaCoverage[] }) {
  const styles = useStyles();
  const [open, setOpen] = useState(false);

  if (coverages.length === 0) {
    return <CoverageTable coverages={coverages} />;
  }

  const sorted = sortCoveragesByAccess(coverages);

  return (
    <>
      <div
        className={mergeClasses(styles.subCaption, styles.subCaptionToggle)}
        role="button"
        tabIndex={0}
        onClick={() => setOpen((o) => !o)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            setOpen((o) => !o);
          }
        }}
      >
        <span className={styles.chevron}>
          {open ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
        </span>
        <ShieldTask16Regular />
        {coverages.length} conditional access {coverages.length === 1 ? 'policy' : 'policies'}
      </div>
      {open && <CoverageTable coverages={sorted} />}
    </>
  );
}

/** Block / grant roll-up for an app header (mirrors IA ProfileSummaryChips). */
function CoverageSummary({ coverages }: { coverages: AppCaCoverage[] }) {
  const styles = useStyles();
  if (coverages.length === 0) {
    return (
      <Badge appearance="ghost" color="subtle">
        No CA
      </Badge>
    );
  }
  const blocked = coverages.filter((c) => c.access === 'block').length;
  const grant = coverages.filter((c) => c.access === 'grant').length;
  return (
    <span className={styles.summaryRow}>
      {blocked > 0 && (
        <Badge appearance="tint" color="danger">
          {blocked} block
        </Badge>
      )}
      {grant > 0 && (
        <Badge appearance="tint" color="success">
          {grant} grant
        </Badge>
      )}
    </span>
  );
}

// ---------------------------------------------------------------------------
// Application segments sub-table (per-app destinations) — ztspecs 25395
// ---------------------------------------------------------------------------

/** Icon + colour for a segment's destination breadth. */
const BREADTH_META: Record<
  SegmentRow['breadth'],
  { label: string; icon: typeof Tag16Regular; color: BadgeProps['color'] }
> = {
  granular: { label: 'Granular', icon: CheckmarkCircle16Filled, color: 'success' },
  broad: { label: 'Broad', icon: Warning16Filled, color: 'warning' },
  wildcard: { label: 'DNS suffix', icon: Globe16Regular, color: 'informative' },
};

function SegmentTable({ segments }: { segments: SegmentRow[] }) {
  const styles = useStyles();
  return (
    <>
      <div className={styles.subCaption}>
        <Server16Regular />
        Application segments ({segments.length})
      </div>
      <div className={styles.gridWrap}>
        <div className={mergeClasses(styles.segGridRow, styles.gridHeaderRow)}>
          <span>Destination</span>
          <span>Type</span>
          <span>Ports</span>
          <span>Protocol</span>
          <span>Scope</span>
        </div>
        {segments.map((s, i) => {
          const b = BREADTH_META[s.breadth];
          const BIcon = b.icon;
          return (
            <div
              key={s.id}
              className={mergeClasses(
                styles.segGridRow,
                styles.gridDataRow,
                i % 2 === 1 && styles.gridRowAlt,
              )}
            >
              <span className={mergeClasses(styles.gridCell, styles.gridCellWrap, styles.mono)}>
                {s.destinationHost}
              </span>
              <span className={styles.gridCell}>
                <span className={styles.cellBadgeGap}>
                  <Badge appearance="outline" color="brand">
                    {s.destinationTypeLabel}
                  </Badge>
                  {s.hasPort53 && (
                    <Tooltip content="Publishes port 53 (DNS over the tunnel)." relationship="label">
                      <Badge appearance="tint" color="informative" icon={<Globe16Regular />}>
                        DNS:53
                      </Badge>
                    </Tooltip>
                  )}
                </span>
              </span>
              <span className={mergeClasses(styles.gridCell, styles.gridCellWrap, styles.mono)}>
                <span className={styles.cellBadgeGap}>
                  {s.ports}
                  {s.widePorts && (
                    <Tooltip
                      content="More than 10 ports opened (AD ephemeral RPC ranges are exempt)."
                      relationship="label"
                    >
                      <Badge appearance="tint" color="warning" icon={<Warning16Filled />}>
                        wide
                      </Badge>
                    </Tooltip>
                  )}
                </span>
              </span>
              <span className={styles.gridCell}>{s.protocol}</span>
              <span className={styles.gridCell}>
                <Badge appearance="tint" color={b.color} icon={<BIcon />} className={styles.cellBadge}>
                  {b.label}
                </Badge>
              </span>
            </div>
          );
        })}
      </div>
    </>
  );
}

// ---------------------------------------------------------------------------
// App group section (one collapsible card per application)
// ---------------------------------------------------------------------------

function AppGroup({
  name,
  appId,
  icon,
  headerBadges,
  coverages,
  warn,
  expanded,
  onToggle,
  extraBody,
}: {
  name: string;
  appId: string;
  icon: ReactElement;
  headerBadges: ReactElement;
  coverages: AppCaCoverage[];
  warn: boolean;
  expanded: boolean;
  onToggle: () => void;
  extraBody?: ReactElement;
}) {
  const styles = useStyles();
  return (
    <section className={styles.group}>
      <div
        className={mergeClasses(styles.header, warn && styles.headerWarn)}
        role="button"
        tabIndex={0}
        onClick={onToggle}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            onToggle();
          }
        }}
      >
        <div className={styles.headerLead}>
          <div className={styles.headerLeadTop}>
            <span className={styles.chevron}>
              {expanded ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
            </span>
            <span className={styles.appIcon}>{icon}</span>
            <span className={styles.appName}>{name}</span>
          </div>
          <div className={styles.headerMeta}>{headerBadges}</div>
        </div>

        <div className={styles.headerVerdict}>
          <CoverageSummary coverages={coverages} />
        </div>

        <div className={styles.headerCaCount}>
          <Badge appearance="ghost" color="subtle">
            {coverages.length} {coverages.length === 1 ? 'CA policy' : 'CA policies'}
          </Badge>
        </div>

        <div className={styles.headerRight}>
          <span className={styles.appId}>{appId}</span>
        </div>
      </div>

      {expanded && (
        <>
          {extraBody}
          <CoverageSection coverages={coverages} />
        </>
      )}
    </section>
  );
}

function PrivateAppHeaderBadges({ row }: { row: PrivateAccessRow }) {
  const styles = useStyles();
  return (
    <>
      <div className={mergeClasses(styles.cell, styles.cellKind)}>
        <Badge appearance="tint" color={row.kind === 'quickAccess' ? 'informative' : 'brand'}>
          {privateAccessKindLabel(row.kind)}
        </Badge>
      </div>
      <div className={mergeClasses(styles.cell, styles.cellSeg)}>
        {row.segments.length > 0 && (
          <Tooltip
            content={`${row.segments.length} application segment${row.segments.length === 1 ? '' : 's'} configured.`}
            relationship="label"
          >
            <Badge appearance="outline" icon={<Server16Regular />}>
              {row.segments.length} {row.segments.length === 1 ? 'segment' : 'segments'}
            </Badge>
          </Tooltip>
        )}
      </div>
      <div className={mergeClasses(styles.cell, styles.cellDns)}>
        {row.dns.configured ? (
          <Tooltip
            content={
              row.dns.resolutionEnabled
                ? `Private DNS enabled${row.dns.suffixCount > 0 ? ` · ${row.dns.suffixCount} suffix${row.dns.suffixCount === 1 ? '' : 'es'}` : ''}${row.dns.hasPort53 ? ' · port 53 published' : ''}.`
                : `DNS reachable via ${row.dns.suffixCount > 0 ? 'DNS-suffix segment' : 'port 53'}.`
            }
            relationship="label"
          >
            <Badge appearance="tint" color="success" icon={<Globe16Regular />}>
              Private DNS
            </Badge>
          </Tooltip>
        ) : (
          <Tooltip
            content="No private DNS: no DNS-resolution flag, DNS-suffix segment, or port-53 segment."
            relationship="label"
          >
            <Badge appearance="outline" color="subtle">
              No private DNS
            </Badge>
          </Tooltip>
        )}
      </div>
      <div className={mergeClasses(styles.cell, styles.cellAssign)}>
        {row.assignment.hasGap ? (
          <Tooltip
            content="Assignment required but no users or groups are assigned — users cannot reach this app."
            relationship="label"
          >
            <Badge appearance="tint" color="danger" icon={<PersonProhibited16Regular />}>
              No assignment
            </Badge>
          </Tooltip>
        ) : row.assignment.required ? (
          <Tooltip
            content={row.assignment.principals.map((p) => `${p.label} (${p.type})`).join(', ') || 'Assigned'}
            relationship="label"
          >
            <Badge appearance="outline" icon={<People16Regular />}>
              {row.assignment.count} assigned
            </Badge>
          </Tooltip>
        ) : (
          <Tooltip content="Assignment not required — all users have implicit access." relationship="label">
            <Badge appearance="outline" color="informative" icon={<People16Regular />}>
              All users
            </Badge>
          </Tooltip>
        )}
      </div>
      <div className={mergeClasses(styles.cell, styles.cellCsa)}>
        {row.hasCustomSecurityAttributes && (
          <Tooltip content="Custom security attributes are assigned to this app." relationship="label">
            <Badge appearance="outline" icon={<Tag16Regular />}>
              CSA
            </Badge>
          </Tooltip>
        )}
      </div>
    </>
  );
}

function AppProxyHeaderBadges({ row }: { row: AppProxyRow }) {
  const styles = useStyles();
  let badge: ReactElement | null;
  if (row.preAuthEnforced) {
    badge = (
      <Badge appearance="tint" color="success" icon={<CheckmarkCircle16Filled />}>
        Entra pre-auth
      </Badge>
    );
  } else if (row.externalAuthenticationType === 'passthru') {
    badge = (
      <Tooltip
        content="Passthrough: the app is reachable without Entra pre-authentication."
        relationship="label"
      >
        <Badge appearance="tint" color="warning" icon={<Warning16Filled />}>
          Passthrough
        </Badge>
      </Tooltip>
    );
  } else {
    badge = row.externalAuthenticationType ? (
      <Badge appearance="outline">{row.externalAuthenticationType}</Badge>
    ) : null;
  }
  // Mirror the Private Access slot layout so the CA / coverage columns align.
  return (
    <>
      <div className={mergeClasses(styles.cell, styles.cellKind)}>{badge}</div>
      <div className={mergeClasses(styles.cell, styles.cellSeg)} />
      <div className={mergeClasses(styles.cell, styles.cellDns)} />
      <div className={mergeClasses(styles.cell, styles.cellAssign)} />
      <div className={mergeClasses(styles.cell, styles.cellCsa)} />
    </>
  );
}

// ---------------------------------------------------------------------------
// Public component
// ---------------------------------------------------------------------------

export function PrivateAccessView({ privateAccess, caDetails, forwardingProfiles = [] }: PrivateAccessViewProps) {
  const styles = useStyles();
  const [search, setSearch] = useState('');
  // Start with all apps collapsed so the user sees the full inventory at a glance.
  const [collapsed, setCollapsed] = useState<Set<string>>(() => {
    const initial = buildPrivateAccessViewModel(
      privateAccess.apps,
      privateAccess.appProxyApps,
      privateAccess.authStrength,
      caDetails,
    );
    return new Set([
      ...initial.apps.map((r) => r.id),
      ...initial.appProxyApps.map((r) => r.id),
    ]);
  });

  const vm = useMemo(
    () =>
      buildPrivateAccessViewModel(
        privateAccess.apps,
        privateAccess.appProxyApps,
        privateAccess.authStrength,
        caDetails,
      ),
    [privateAccess, caDetails],
  );

  const q = search.trim().toLowerCase();
  const filtering = q.length > 0;
  const apps = filtering ? vm.apps.filter((r) => r.search.includes(q)) : vm.apps;
  const proxyApps = filtering
    ? vm.appProxyApps.filter((r) => r.search.includes(q))
    : vm.appProxyApps;

  const totalApps = vm.apps.length + vm.appProxyApps.length;
  const visibleApps = apps.length + proxyApps.length;

  // Group PA apps by forwarding profile when private profiles are available.
  // Falls back to a flat list when forwardingProfiles is empty (V0 / old fixtures).
  const paProfiles = forwardingProfiles
    .filter((p) => p.trafficForwardingType === 'private')
    .sort((a, b) => (a.priority ?? 999) - (b.priority ?? 999));

  const paGroups = paProfiles.map((profile) => {
    const profileIds = profile.linkedApps
      ? new Set(profile.linkedApps.map((a) => a.appId))
      : null; // null = no linkedApps → all apps (graceful fallback for old fixtures)
    const rows = profileIds ? apps.filter((r) => profileIds.has(r.appId)) : [...apps];
    return { profile, rows };
  });
  const paAssignedIds = new Set(paGroups.flatMap(({ rows }) => rows.map((r) => r.appId)));
  const paUnassigned = paProfiles.length > 0
    ? apps.filter((r) => !paAssignedIds.has(r.appId))
    : [];

  const isExpanded = (id: string) => filtering || !collapsed.has(id);
  const toggle = (id: string) =>
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  const allIds = useMemo(
    () => [...vm.apps.map((r) => r.id), ...vm.appProxyApps.map((r) => r.id)],
    [vm],
  );
  const expandAll = () => setCollapsed(new Set());
  const collapseAll = () => setCollapsed(new Set(allIds));

  if (totalApps === 0) {
    return (
      <div className={styles.empty}>
        <LockClosed20Regular style={{ fontSize: '32px' }} />
        <Text weight="semibold">No Private Access data in this data source.</Text>
        <Text>
          Private Access apps, Application Proxy apps, and their Conditional Access
          correlation are captured by the test-harness exporter. Load an exported
          fixture that includes the <code>privateAccessApps</code> /{' '}
          <code>appProxyApps</code> arrays to inspect them here.
        </Text>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.toolbar}>
        <Input
          className={styles.search}
          value={search}
          onChange={(_, d) => setSearch(d.value)}
          contentBefore={<Search16Regular />}
          placeholder="Search apps, app IDs, URLs…"
        />
        <Text size={200} className={styles.count}>
          {visibleApps} of {totalApps} apps · {vm.coveredAppCount} CA-covered
        </Text>
        <span className={styles.grow} />
        <Button
          size="small"
          appearance="subtle"
          icon={<Open16Regular />}
          disabled={filtering}
          onClick={expandAll}
        >
          Expand all
        </Button>
        <Button size="small" appearance="subtle" disabled={filtering} onClick={collapseAll}>
          Collapse all
        </Button>
      </div>

      <div className={styles.scroll}>
        {/* PA apps: grouped by profile when private forwarding profiles exist, flat otherwise */}
        {paGroups.length === 0 && apps.length > 0 && (
          <section className={styles.section}>
            <div className={styles.sectionHead}>
              <LockClosed20Regular />
              <Text className={styles.sectionTitle}>Private Access applications</Text>
              <Badge appearance="ghost" color="subtle">{apps.length}</Badge>
            </div>
            {apps.map((row) => (
              <AppGroup
                key={row.id}
                name={row.name}
                appId={row.appId || row.id}
                icon={row.kind === 'quickAccess' ? <Globe20Regular /> : <LockClosed20Regular />}
                headerBadges={<PrivateAppHeaderBadges row={row} />}
                coverages={row.coverages}
                warn={row.assignment.hasGap}
                expanded={isExpanded(row.id)}
                onToggle={() => toggle(row.id)}
                extraBody={
                  row.segments.length > 0 ? <SegmentTable segments={row.segments} /> : undefined
                }
              />
            ))}
          </section>
        )}
        {paGroups.length > 0 && paGroups.map(({ profile, rows }) => {
          if (rows.length === 0 && filtering) return null;
          const assocCount   = profile.associations?.length ?? 0;
          const assignedLabel = assocCount === 0 ? 'All users and devices' : `${assocCount} assignment(s) configured`;
          const isFailClose  = profile.clientFallbackAction === 'block';
          return (
            <section key={profile.id ?? profile.name} className={styles.section}>
              <div className={styles.profileHeadRow}>
                <div className={styles.sectionHead}>
                  <LockClosed20Regular />
                  <Text className={styles.sectionTitle}>{profile.name ?? 'Private Access'}</Text>
                  <Badge appearance="ghost" color="subtle">
                    {rows.length} app{rows.length !== 1 ? 's' : ''}
                  </Badge>
                  {profile.state === 'enabled' && (
                    <Badge appearance="tint" color="success" icon={<CheckmarkCircle16Filled />}>on</Badge>
                  )}
                  {profile.state === 'disabled' && (
                    <Badge appearance="tint" color="warning" icon={<Warning16Filled />}>off</Badge>
                  )}
                  {/* Read-only toggle — pointerEvents:none keeps full button styling without disabled graying.
                      Remove the wrapper and add onClick to enable writes in V1+. */}
                  <div style={{ pointerEvents: 'none' }}>
                    <ToggleButton
                      size="small"
                      appearance="outline"
                      checked={isFailClose}
                      icon={isFailClose ? <LockClosed16Regular /> : undefined}
                    >
                      {isFailClose ? 'Fail-close' : 'Fail-open'}
                    </ToggleButton>
                  </div>
                  {profile.priority !== undefined && (
                    <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                      Priority {profile.priority}
                    </Text>
                  )}
                </div>
                <div className={styles.profileMeta}>
                  <span>Assigned to: {assignedLabel}</span>
                  {isFailClose ? (
                    <span style={{ color: tokens.colorPaletteMarigoldForeground2 }}>
                      Traffic is dropped if the GSA edge is unreachable.
                    </span>
                  ) : (
                    <span>Traffic is allowed direct if the GSA edge is unreachable.</span>
                  )}
                </div>
              </div>
              {rows.length === 0
                ? <div className={styles.noMatches}>No apps in this profile.</div>
                : rows.map((row) => (
                    <AppGroup
                      key={row.id}
                      name={row.name}
                      appId={row.appId || row.id}
                      icon={row.kind === 'quickAccess' ? <Globe20Regular /> : <LockClosed20Regular />}
                      headerBadges={<PrivateAppHeaderBadges row={row} />}
                      coverages={row.coverages}
                      warn={row.assignment.hasGap}
                      expanded={isExpanded(row.id)}
                      onToggle={() => toggle(row.id)}
                      extraBody={
                        row.segments.length > 0 ? <SegmentTable segments={row.segments} /> : undefined
                      }
                    />
                  ))
              }
            </section>
          );
        })}
        {paUnassigned.length > 0 && (
          <section className={styles.section}>
            <div className={styles.sectionHead}>
              <LockClosed20Regular />
              <Text className={styles.sectionTitle}>Not assigned to a profile</Text>
              <Badge appearance="ghost" color="subtle">{paUnassigned.length}</Badge>
            </div>
            {paUnassigned.map((row) => (
              <AppGroup
                key={row.id}
                name={row.name}
                appId={row.appId || row.id}
                icon={row.kind === 'quickAccess' ? <Globe20Regular /> : <LockClosed20Regular />}
                headerBadges={<PrivateAppHeaderBadges row={row} />}
                coverages={row.coverages}
                warn={row.assignment.hasGap}
                expanded={isExpanded(row.id)}
                onToggle={() => toggle(row.id)}
                extraBody={
                  row.segments.length > 0 ? <SegmentTable segments={row.segments} /> : undefined
                }
              />
            ))}
          </section>
        )}

        {proxyApps.length > 0 && (
          <section className={styles.section}>
            <div className={styles.sectionHead}>
              <AppGeneric20Regular />
              <Text className={styles.sectionTitle}>Application Proxy applications</Text>
              <Badge appearance="ghost" color="subtle">
                {proxyApps.length}
              </Badge>
            </div>
            {proxyApps.map((row) => (
              <AppGroup
                key={row.id}
                name={row.name}
                appId={row.appId || row.id}
                icon={<AppGeneric20Regular />}
                headerBadges={<AppProxyHeaderBadges row={row} />}
                coverages={row.coverages}
                warn={row.externalAuthenticationType === 'passthru'}
                expanded={isExpanded(row.id)}
                onToggle={() => toggle(row.id)}
              />
            ))}
          </section>
        )}

        {visibleApps === 0 && (
          <div className={styles.noMatches}>No apps match “{search}”.</div>
        )}
      </div>
    </div>
  );
}
