/**
 * Unified policy table — Tier 1 (UI).
 *
 * Renders the policy tree grouped into collapsible Security Profile sections.
 * Each profile header shows its priority, policy/rule counts, and the linked
 * Conditional Access policies (with user/group targeting when CA detail is
 * available — file-mode fixtures or tenants where Policy.Read.All was granted;
 * otherwise it degrades gracefully).
 *
 * TanStack Table (headless) powers the in-memory search + sort over the leaf
 * policy/rule rows; we do NOT use Fluent `DataGrid` per the architecture's
 * table rule. Grouping is applied over the filtered/sorted leaf rows so that
 * profiles stay contiguous and only matching groups remain visible.
 *
 * Virtualisation (TanStack Virtual) is deferred — fixture sets are small. Add
 * it before shipping tenants expected to exceed ~200 rules.
 */

import { useMemo, useState, createContext, useContext, type ReactNode } from 'react';
import {
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  useReactTable,
  type ColumnDef,
  type SortingState,
  type FilterFn,
} from '@tanstack/react-table';
import {
  Badge,
  Button,
  Dropdown,
  Input,
  Option,
  Text,
  Tooltip,
  OverlayDrawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  makeStyles,
  mergeClasses,
  tokens,
  type BadgeProps,
} from '@fluentui/react-components';
import {
  ChevronDown16Regular,
  ChevronRight16Regular,
  Search16Regular,
  ShieldKeyhole16Regular,
  ShieldCheckmark16Regular,
  Person16Regular,
  PeopleTeam16Regular,
  Globe16Regular,
  Link16Regular,
  Tag16Regular,
  Server16Regular,
  Chat16Regular,
  Prohibited16Regular,
  Open16Regular,
  ToggleRight16Filled,
  ToggleLeft16Filled,
  Dismiss24Regular,
} from '@fluentui/react-icons';
import type {
  LinkedCaInfo,
  PolicyRow,
  PrincipalKind,
  PrincipalRef,
  ProfileGroup,
} from './policyRows';
import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type { FilteringPolicyLink } from '../../model/definitions/WebContentFilteringPolicy.definition';
import { resolvePolicyTypeLabel } from '../../model/registry';
import { actionColor } from '../actionColor';
import {
  POLICY_TYPE_COLOR,
  shortPolicyType,
} from '../policyTypeMeta';
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
    gap: tokens.spacingVerticalM,
  },

  // --- Profile group card ---------------------------------------------------
  group: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusXLarge,
    overflow: 'hidden',
    boxShadow: tokens.shadow4,
    backgroundColor: tokens.colorNeutralBackground1,
    // Prevent the card from being compressed below its content height inside
    // the scrollable flex column (which would let an expanded sub-table bleed
    // over the next profile card).
    flexShrink: 0,
  },
  // Fixed-column swim-lane grid so every badge group lands in the same column
  // on every profile card — the collapsed list reads like a table (item 6).
  // Lanes: priority gutter | identity (name + counts subtext) | action summary
  // | conditional access. Alignment only, no divider lines.
  //
  // Identity and action-summary are FIXED px (not `minmax(0, …fr)`) so this
  // header's "before the verdict lane" offset (64 + gap + 440 + gap = 528px)
  // is a constant, independent of window width — and it deliberately matches
  // PrivateAccessView.tsx's `.header` cumulative offset before its verdict
  // lane (also 528px), so the "blocked/allowed" (IA) and "block/grant" (PA)
  // lanes land at the same x on the page, at any window size. Only the
  // trailing CA lane stays flexible.
  header: {
    display: 'grid',
    gridTemplateColumns:
      '64px 440px 200px minmax(0, 1fr)',
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
  headerBaseline: {
    borderLeftColor: tokens.colorPaletteGreenBorderActive,
    background: `linear-gradient(90deg, ${tokens.colorPaletteGreenBackground2} 0%, ${tokens.colorNeutralBackground1} 70%)`,
    ':hover': {
      background: `linear-gradient(90deg, ${tokens.colorPaletteGreenBackground2} 0%, ${tokens.colorNeutralBackground1Hover} 70%)`,
    },
  },
  // --- Lane 0: priority gutter ----------------------------------------------
  gutter: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start',
    justifyContent: 'center',
    lineHeight: tokens.lineHeightBase200,
    minWidth: 0,
  },
  gutterLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
  },
  gutterValue: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  // --- Lane 1: identity (name + type badge + counts subtext) ----------------
  identity: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    minWidth: 0,
  },
  identityLead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  chevron: {
    display: 'inline-flex',
    color: tokens.colorNeutralForeground2,
    flexShrink: 0,
  },
  profileName: {
    fontWeight: tokens.fontWeightBold,
    fontSize: tokens.fontSizeBase400,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    minWidth: 0,
  },
  counts: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: 'nowrap',
    paddingLeft: '24px',
  },
  // --- Lane 2: action summary -----------------------------------------------
  actionLane: {
    display: 'flex',
    alignItems: 'center',
    minWidth: 0,
  },
  // --- Lane 3: conditional access -------------------------------------------
  caLane: {
    display: 'flex',
    alignItems: 'center',
    minWidth: 0,
  },
  caGroup: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
    justifyContent: 'flex-start',
  },
  caLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: 'nowrap',
  },

  // --- header targeting (CA + principals) -----------------------------------
  targetsWrap: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start',
    gap: tokens.spacingVerticalXS,
    minWidth: 0,
  },
  principalRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
    justifyContent: 'flex-start',
  },
  peopleIcon: {
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  excludeChip: {
    textDecorationLine: 'line-through',
    opacity: 0.85,
  },
  chipLabel: {
    display: 'inline-block',
    maxWidth: '180px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    verticalAlign: 'bottom',
  },

  // --- toolbar filters ------------------------------------------------------
  filterDropdown: { minWidth: '130px' },
  ruleFilter: { width: '160px' },
  typeCell: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
  },

  // --- Inner rule rows ------------------------------------------------------
  muted: { color: tokens.colorNeutralForeground3 },
  destinations: {
    width: '100%',
    color: tokens.colorNeutralForeground2,
    display: '-webkit-box',
    WebkitLineClamp: 2,
    WebkitBoxOrient: 'vertical',
    overflow: 'hidden',
    wordBreak: 'break-word',
  },
  destinationLines: {
    width: '100%',
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    color: tokens.colorNeutralForeground2,
  },
  destinationLine: {
    wordBreak: 'break-word',
  },
  mark: {
    backgroundColor: tokens.colorPaletteYellowBackground2,
    color: tokens.colorNeutralForeground1,
    borderRadius: tokens.borderRadiusSmall,
    padding: '0 1px',
  },

  // --- details drawer -------------------------------------------------------
  detailHeaderTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  detailBody: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalXL,
  },
  detailGrid: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
  detailRow: {
    display: 'grid',
    gridTemplateColumns: '160px 1fr',
    gap: tokens.spacingHorizontalM,
    alignItems: 'start',
  },
  detailKey: {
    color: tokens.colorNeutralForeground3,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
  },
  detailValue: {
    color: tokens.colorNeutralForeground1,
    wordBreak: 'break-word',
  },
  detailJsonLabel: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    marginTop: tokens.spacingVerticalS,
  },
  detailJson: {
    margin: 0,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    overflowX: 'auto',
  },
  cellBadge: {
    whiteSpace: 'nowrap',
  },
  emptyProfile: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },  noMatches: {
    padding: tokens.spacingVerticalXXL,
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },

  // --- collapsed profile summary --------------------------------------------
  summaryRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
  },

  // --- nested policy sub-cards (shared fixed-column grid) -------------------
  // One grid template drives the column-title row, every policy sub-card
  // header, and every rule row, so Action (and every other column) lines up
  // vertically across all policies in a profile — the same rigor as the
  // profile cards (item 6). Lanes: name | type | action | rule-type | dest.
  policyList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalM,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  gridRow: {
    display: 'grid',
    gridTemplateColumns:
      'minmax(0, 1.5fr) 148px 108px 148px minmax(0, 1.9fr)',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalM,
  },
  colTitles: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    color: tokens.colorNeutralForeground3,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.03em',
  },
  subCard: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusLarge,
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  subHeader: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  ruleRow: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderTop: `1px solid ${tokens.colorNeutralStroke3}`,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  ruleRowAlt: { backgroundColor: tokens.colorNeutralBackground2 },
  gridCell: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
  },
  nameLead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
  },
  stateIcon: {
    display: 'inline-flex',
    alignItems: 'center',
    flexShrink: 0,
    fontSize: '16px',
    lineHeight: 0,
  },
  stateEnabled: { color: tokens.colorPaletteGreenForeground1 },
  stateDisabled: { color: tokens.colorPaletteDarkOrangeForeground1 },
  subName: {
    fontWeight: tokens.fontWeightSemibold,
    minWidth: 0,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  ruleNameCell: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
    paddingLeft: '24px',
  },
  ruleNameText: {
    minWidth: 0,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  rulePriorityTag: {
    flexShrink: 0,
    minWidth: '20px',
    textAlign: 'center',
    padding: `0 ${tokens.spacingHorizontalXXS}`,
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    fontFamily: tokens.fontFamilyMonospace,
  },
  ruleNamePlaceholder: {
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic',
    paddingLeft: '24px',
  },

  // --- targeting tooltip ----------------------------------------------------
  tip: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    maxWidth: '320px',
    maxHeight: '50vh',
    overflowY: 'auto',
  },
  tipLabel: { fontWeight: tokens.fontWeightSemibold },
  tipValue: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    color: tokens.colorNeutralForeground2,
    wordBreak: 'break-word',
  },
  tipLine: { wordBreak: 'break-word' },
});

/** Icon + colour for each rule type, shown in the Type column. */
const RULE_TYPE_META: Record<
  string,
  { icon: typeof Globe16Regular; color: BadgeProps['color'] }
> = {
  webCategory: { icon: Tag16Regular, color: 'brand' },
  fqdn: { icon: Globe16Regular, color: 'informative' },
  url: { icon: Link16Regular, color: 'success' },
  ipAddress: { icon: Server16Regular, color: 'warning' },
  ipSubnet: { icon: Server16Regular, color: 'warning' },
  ipRange: { icon: Server16Regular, color: 'warning' },
  prompt: { icon: Chat16Regular, color: 'success' },
};

/**
 * Search-term highlight context. Provides the active search keyword to nested
 * cells so they can emphasise matching substrings without prop drilling.
 */
const SearchHighlightContext = createContext('');

/** Render `text` with any occurrence of the active search keyword highlighted. */
function Highlight({ text }: { text: string }): ReactNode {
  const query = useContext(SearchHighlightContext).trim();
  const styles = useStyles();
  if (query === '' || text === '') return text;

  const lower = text.toLowerCase();
  const needle = query.toLowerCase();
  if (!lower.includes(needle)) return text;

  const parts: ReactNode[] = [];
  let i = 0;
  let idx = lower.indexOf(needle, i);
  let key = 0;
  while (idx !== -1) {
    if (idx > i) parts.push(text.slice(i, idx));
    parts.push(
      <mark key={key++} className={styles.mark}>
        {text.slice(idx, idx + query.length)}
      </mark>,
    );
    i = idx + query.length;
    idx = lower.indexOf(needle, i);
  }
  if (i < text.length) parts.push(text.slice(i));
  return <>{parts}</>;
}

/**
 * Double-click → "Details" drawer wiring. Provides openers for a profile,
 * policy, or rule so deeply nested headers/rows can request the details panel
 * without prop drilling. Null when no table is mounting the provider.
 */
interface DetailOpeners {
  openProfile: (profileId: string) => void;
  openPolicy: (policyKey: string) => void;
  openRule: (rowKey: string) => void;
}
const DetailContext = createContext<DetailOpeners | null>(null);
function useDetailOpeners(): DetailOpeners | null {
  return useContext(DetailContext);
}

/** Suppress the browser's text selection that a double-click would otherwise make. */
function clearSelection() {
  window.getSelection?.()?.removeAllRanges();
}

/** Icon for a resolved CA principal, by kind. */
function principalIcon(kind: PrincipalKind) {
  switch (kind) {
    case 'group':
      return PeopleTeam16Regular;
    case 'all':
      return Globe16Regular;
    case 'none':
      return Prohibited16Regular;
    case 'guests':
    case 'user':
    case 'role':
    case 'other':
    default:
      return Person16Regular;
  }
}

/** Composite filter value driving the single TanStack global filter. */
interface TableFilters {
  search: string;
  action: string;
  ruleType: string;
  rule: string;
}

const EMPTY_FILTERS: TableFilters = { search: '', action: '', ruleType: '', rule: '' };

function filtersActive(f: TableFilters): boolean {
  return Boolean(f.search || f.action || f.ruleType || f.rule);
}

/** Match leaf rows against the search box + the per-column filters. */
function makeGlobalFilter(
  groupById: Map<string, ProfileGroup>,
): FilterFn<PolicyRow> {
  return (row, _columnId, filterValue: TableFilters) => {
    const r = row.original;

    if (filterValue.action && r.action !== filterValue.action) return false;
    if (filterValue.ruleType && r.ruleType !== filterValue.ruleType) return false;
    if (
      filterValue.rule &&
      !r.ruleName.toLowerCase().includes(filterValue.rule.toLowerCase())
    ) {
      return false;
    }

    const needle = filterValue.search.toLowerCase().trim();
    if (!needle) return true;
    const group = groupById.get(r.profileId);
    const haystack = [
      r.policyName,
      r.policyType,
      r.ruleName,
      r.ruleType,
      r.destinations,
      group?.profileName ?? r.profileName,
      group?.caSearchText ?? r.caPolicyNames,
    ]
      .join(' ')
      .toLowerCase();
    return haystack.includes(needle);
  };
}

// ---------------------------------------------------------------------------
// CA targeting chip + resolved principals
// ---------------------------------------------------------------------------

function CaChip({ ca }: { ca: LinkedCaInfo }) {
  const styles = useStyles();
  const Icon = ca.hasDetail ? ShieldCheckmark16Regular : ShieldKeyhole16Regular;

  const tip = ca.hasDetail ? (
    ca.targets.length > 0 ? (
      <div className={styles.tip}>
        {ca.targets.map((t) => (
          <div key={t.label}>
            <Text className={styles.tipLabel}>
              {t.label} ({t.principals.length})
            </Text>
            <div className={styles.tipValue}>
              {t.principals.map((p) => (
                <span key={`${p.mode}/${p.id}`} className={styles.tipLine}>
                  {p.label}
                </span>
              ))}
            </div>
          </div>
        ))}
      </div>
    ) : (
      <Text>No user/group conditions on this policy.</Text>
    )
  ) : (
    <Text>
      Targeting detail unavailable — Policy.Read.All not granted (spec §2.3).
    </Text>
  );

  return (
    <Tooltip content={tip} relationship="label" withArrow>
      <Badge
        appearance="tint"
        color={ca.hasDetail ? 'brand' : 'subtle'}
        icon={<Icon />}
        style={{ maxWidth: '240px' }}
      >
        <span
          style={{
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
          }}
        >
          {ca.name}
        </span>
      </Badge>
    </Tooltip>
  );
}

function PrincipalChip({ principal }: { principal: PrincipalRef }) {
  const styles = useStyles();
  const Icon = principalIcon(principal.kind);
  const isExclude = principal.mode === 'exclude';
  const tip = `${isExclude ? 'Excluded' : 'Included'} ${principal.kind} · ${principal.id}`;

  return (
    <Tooltip content={tip} relationship="label" withArrow>
      <Badge
        appearance={isExclude ? 'outline' : 'tint'}
        color={isExclude ? 'danger' : 'brand'}
        icon={<Icon />}
      >
        <span
          className={mergeClasses(
            styles.chipLabel,
            isExclude && styles.excludeChip,
          )}
        >
          {principal.label}
        </span>
      </Badge>
    </Tooltip>
  );
}

/** How many principal chips to show inline before collapsing into "+N". */
const PRINCIPAL_LIMIT = 6;

function ProfileTargets({ group }: { group: ProfileGroup }) {
  const styles = useStyles();

  if (group.caPolicies.length === 0) {
    return (
      <span className={styles.caLabel}>No Conditional Access policy linked</span>
    );
  }

  const all = [...group.principals.include, ...group.principals.exclude];
  const shown = all.slice(0, PRINCIPAL_LIMIT);
  const overflow = all.slice(PRINCIPAL_LIMIT);

  return (
    <div className={styles.targetsWrap}>
      <div className={styles.caGroup}>
        {group.caPolicies.map((ca) => (
          <CaChip key={ca.id} ca={ca} />
        ))}
      </div>

      <div className={styles.principalRow}>
        {all.length === 0 ? (
          <span className={styles.caLabel}>
            {group.hasCaDetail
              ? 'No user / group scope'
              : 'Targeting detail unavailable'}
          </span>
        ) : (
          <>
            {shown.map((p) => (
              <PrincipalChip key={`${p.mode}/${p.id}`} principal={p} />
            ))}
            {overflow.length > 0 && (
              <Tooltip
                relationship="label"
                withArrow
                content={
                  <div className={styles.tip}>
                    {overflow.map((p) => (
                      <Text key={`${p.mode}/${p.id}`}>
                        {p.mode === 'exclude' ? '− ' : '+ '}
                        {p.label}
                      </Text>
                    ))}
                  </div>
                }
              >
                <Badge appearance="ghost" color="subtle">
                  +{overflow.length} more
                </Badge>
              </Tooltip>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Shared-grid cell renderers + rule rows
// ---------------------------------------------------------------------------

/** Column titles for the shared expanded-level grid. */
const POLICY_GRID_TITLES = [
  'Policy Name',
  'Policy Type',
  'Action',
  'Rule Type',
  'Destinations',
];

function PolicyTypeBadge({
  label,
  unknown,
}: {
  label: string;
  unknown: boolean;
}) {
  const styles = useStyles();
  if (label === '') return <span className={styles.muted}>—</span>;
  if (unknown) return <span className={styles.muted}>{shortPolicyType(label)}</span>;
  return (
    <Badge
      appearance="filled"
      color={POLICY_TYPE_COLOR[label] ?? 'brand'}
      className={styles.cellBadge}
    >
      {shortPolicyType(label)}
    </Badge>
  );
}

function ActionBadge({ action }: { action: string }) {
  const styles = useStyles();
  if (action === '') return <span className={styles.muted}>—</span>;
  return (
    <Badge appearance="filled" color={actionColor(action)} className={styles.cellBadge}>
      {action}
    </Badge>
  );
}

function RuleTypeBadge({ ruleType }: { ruleType: string }) {
  const styles = useStyles();
  if (ruleType === '') return <span className={styles.muted}>—</span>;
  const meta = RULE_TYPE_META[ruleType];
  const Icon = meta?.icon ?? Tag16Regular;
  return (
    <Badge
      appearance="tint"
      color={meta?.color ?? 'subtle'}
      icon={<Icon />}
      className={styles.cellBadge}
    >
      {ruleType}
    </Badge>
  );
}

function DestinationsCell({ value }: { value: string }) {
  const styles = useStyles();
  if (value === '') return <span className={styles.muted}>—</span>;
  return (
    <Tooltip content={value} relationship="description" withArrow>
      <span className={styles.destinations}>
        <Highlight text={value} />
      </span>
    </Tooltip>
  );
}

/** Destinations rendered one entry per line (expanded rule rows). */
function DestinationsList({ entries, fallback }: { entries: string[]; fallback: string }) {
  const styles = useStyles();
  if (entries.length === 0) {
    if (fallback === '') return <span className={styles.muted}>—</span>;
    return (
      <span className={styles.destinationLine}>
        <Highlight text={fallback} />
      </span>
    );
  }
  return (
    <div className={styles.destinationLines}>
      {entries.map((entry, i) => (
        <span key={i} className={styles.destinationLine}>
          <Highlight text={entry} />
        </span>
      ))}
    </div>
  );
}

/** Rule rows for one policy, on the shared grid (Action aligns with the header). */
function RuleRows({ rows }: { rows: PolicyRow[] }) {
  const styles = useStyles();
  const openers = useDetailOpeners();
  const realRows = rows.filter((r) => r.ruleName !== '');

  if (realRows.length === 0) {
    return (
      <div className={mergeClasses(styles.gridRow, styles.ruleRow)}>
        <span className={styles.ruleNamePlaceholder}>This policy has no rules.</span>
        <span />
        <span />
        <span />
        <span />
      </div>
    );
  }

  return (
    <>
      {realRows.map((r, i) => {
        // TLS inspection rules carry their action per rule; other rows fall
        // back to the policy-level action.
        const actionValue =
          r.ruleAction && r.ruleAction !== '' ? r.ruleAction : r.action;
        return (
          <div
            key={r.key}
            className={mergeClasses(
              styles.gridRow,
              styles.ruleRow,
              i % 2 === 1 && styles.ruleRowAlt,
            )}
            onDoubleClick={(e) => {
              e.preventDefault();
              clearSelection();
              openers?.openRule(r.key);
            }}
            title="Double-click for details"
          >
            <span className={styles.ruleNameCell}>
              {r.rulePriority !== null && (
                <Tooltip
                  content={`Rule priority ${r.rulePriority} (lower is evaluated first)`}
                  relationship="label"
                >
                  <span className={styles.rulePriorityTag}>{r.rulePriority}</span>
                </Tooltip>
              )}
              <span className={styles.ruleNameText}>
                <Highlight text={r.ruleName} />
              </span>
            </span>
            <span />
            <span className={styles.gridCell}>
              <ActionBadge action={actionValue} />
            </span>
            <span className={styles.gridCell}>
              <RuleTypeBadge ruleType={r.ruleType} />
            </span>
            <DestinationsList entries={r.destinationList} fallback={r.destinations} />
          </div>
        );
      })}
    </>
  );
}

// ---------------------------------------------------------------------------
// Nested policy sub-cards + effective-rules summary
// ---------------------------------------------------------------------------

interface PolicyBucket {
  policyKey: string;
  policyName: string;
  policyType: string;
  policyTypeUnknown: boolean;
  action: string;
  policyState: string;
  rows: PolicyRow[];
}

/** Group a profile's flat rows by their owning policy, preserving order. */
function groupRowsByPolicy(rows: PolicyRow[]): PolicyBucket[] {
  const out: PolicyBucket[] = [];
  const indexByKey = new Map<string, number>();
  for (const r of rows) {
    if (r.policyKey === '') continue;
    let idx = indexByKey.get(r.policyKey);
    if (idx === undefined) {
      idx = out.length;
      indexByKey.set(r.policyKey, idx);
      out.push({
        policyKey: r.policyKey,
        policyName: r.policyName,
        policyType: r.policyType,
        policyTypeUnknown: r.policyTypeUnknown,
        action: r.action,
        policyState: r.policyState,
        rows: [],
      });
    }
    out[idx].rows.push(r);
  }
  return out;
}

/** Short, deduplicated preview of a policy's destinations for the collapsed row. */
function destinationsPreview(rows: PolicyRow[]): string {
  const seen = new Set<string>();
  for (const r of rows) {
    if (r.destinations) seen.add(r.destinations);
  }
  return [...seen].join(' · ');
}

/** Block / allow / catch-all roll-up shown on the profile header. */
function ProfileSummaryChips({ group }: { group: ProfileGroup }) {
  const styles = useStyles();
  const { blockRuleCount, allowRuleCount, catchAllAction } = group.summary;
  if (blockRuleCount === 0 && allowRuleCount === 0 && catchAllAction === null) {
    return null;
  }
  return (
    <span className={styles.summaryRow}>
      {blockRuleCount > 0 && (
        <Badge appearance="tint" color="danger">
          {blockRuleCount} blocked
        </Badge>
      )}
      {allowRuleCount > 0 && (
        <Badge appearance="tint" color="success">
          {allowRuleCount} allowed
        </Badge>
      )}
      {catchAllAction && (
        <Tooltip
          content="Default action for traffic not matched by any earlier rule (catch-all)."
          relationship="label"
        >
          <Badge
            appearance="outline"
            color={catchAllAction === 'block' ? 'danger' : 'success'}
          >
            default: {catchAllAction}
          </Badge>
        </Tooltip>
      )}
    </span>
  );
}

/**
 * Enabled / disabled state of a policy link within the profile, shown as a
 * compact toggle icon *before* the policy name (item: state). Green toggle-on
 * = enabled; amber toggle-off = disabled (linked but not enforced). Empty
 * state renders nothing. Icon-only — the tooltip carries the accessible name.
 */
function PolicyStateIcon({ state }: { state: string }) {
  const styles = useStyles();
  if (state === '') return null;
  if (state === 'disabled') {
    return (
      <Tooltip
        content="Disabled — this policy is linked to the profile but its rules are not enforced."
        relationship="label"
      >
        <span
          className={mergeClasses(styles.stateIcon, styles.stateDisabled)}
          aria-label="Disabled"
          role="img"
        >
          <ToggleLeft16Filled />
        </span>
      </Tooltip>
    );
  }
  return (
    <Tooltip content="Enabled — this policy is enforced." relationship="label">
      <span
        className={mergeClasses(styles.stateIcon, styles.stateEnabled)}
        aria-label="Enabled"
        role="img"
      >
        <ToggleRight16Filled />
      </span>
    </Tooltip>
  );
}

function PolicySubCard({
  bucket,
  forceExpand,
}: {
  bucket: PolicyBucket;
  forceExpand: boolean;
}) {
  const styles = useStyles();
  const [open, setOpen] = useState(false);
  const expanded = forceExpand || open;
  const ruleCount = bucket.rows.filter((r) => r.ruleName !== '').length;
  const preview = destinationsPreview(bucket.rows);
  const openers = useDetailOpeners();

  return (
    <section className={styles.subCard}>
      <div
        className={mergeClasses(styles.gridRow, styles.subHeader)}
        role="button"
        tabIndex={0}
        onClick={() => setOpen((v) => !v)}
        onDoubleClick={(e) => {
          e.preventDefault();
          clearSelection();
          openers?.openPolicy(bucket.policyKey);
        }}
        title="Double-click for details"
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            setOpen((v) => !v);
          }
        }}
      >
        <span className={styles.nameLead}>
          <span className={styles.chevron}>
            {expanded ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
          </span>
          <PolicyStateIcon state={bucket.policyState} />
          <span className={styles.subName}>
            <Highlight text={bucket.policyName || '(unnamed policy)'} />
          </span>
          <Badge appearance="ghost" color="subtle">
            {ruleCount} {ruleCount === 1 ? 'rule' : 'rules'}
          </Badge>
        </span>
        <span className={styles.gridCell}>
          <PolicyTypeBadge
            label={bucket.policyType}
            unknown={bucket.policyTypeUnknown}
          />
        </span>
        <span className={styles.gridCell}>
          <ActionBadge action={bucket.action} />
        </span>
        <span />
        {!expanded && preview ? (
          <DestinationsCell value={preview} />
        ) : (
          <span />
        )}
      </div>
      {expanded && <RuleRows rows={bucket.rows} />}
    </section>
  );
}

/**
 * Profile body: every filtering policy is rendered as its own collapsible
 * sub-card so the columns line up identically whether a profile has one policy
 * or many. A column-title row sits above the sub-cards; the sub-card header,
 * rule rows, and titles all share one fixed grid template so Action (and every
 * other column) aligns vertically across the whole profile.
 */
function ProfileBody({
  rows,
  forceExpand,
}: {
  rows: PolicyRow[];
  forceExpand: boolean;
}) {
  const styles = useStyles();
  const buckets = groupRowsByPolicy(rows);

  if (buckets.length === 0) {
    return (
      <div className={styles.emptyProfile}>
        This profile has no filtering policies.
      </div>
    );
  }

  return (
    <div className={styles.policyList}>
      <div className={mergeClasses(styles.gridRow, styles.colTitles)}>
        {POLICY_GRID_TITLES.map((title) => (
          <span key={title}>{title}</span>
        ))}
      </div>
      {buckets.map((bucket) => (
        <PolicySubCard
          key={bucket.policyKey}
          bucket={bucket}
          forceExpand={forceExpand}
        />
      ))}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Profile group section
// ---------------------------------------------------------------------------

function ProfileSection({
  group,
  rows,
  expanded,
  forceExpandPolicies,
  onToggle,
}: {
  group: ProfileGroup;
  rows: PolicyRow[];
  expanded: boolean;
  forceExpandPolicies: boolean;
  onToggle: () => void;
}) {
  const styles = useStyles();
  const ruleCount = rows.filter((r) => r.ruleName !== '').length;
  const openers = useDetailOpeners();

  return (
    <section className={styles.group}>
      <div
        className={mergeClasses(
          styles.header,
          group.isBaseline && styles.headerBaseline,
        )}
        onClick={onToggle}
        onDoubleClick={(e) => {
          e.preventDefault();
          clearSelection();
          openers?.openProfile(group.profileId);
        }}
        title="Double-click for details"
        role="button"
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            onToggle();
          }
        }}
      >
        <div className={styles.gutter}>
          {group.profilePriority !== null && (
            <>
              <span className={styles.gutterLabel}>Priority</span>
              <span className={styles.gutterValue}>{group.profilePriority}</span>
            </>
          )}
        </div>

        <div className={styles.identity}>
          <div className={styles.identityLead}>
            <span className={styles.chevron}>
              {expanded ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
            </span>
            <span className={styles.profileName}>
              <Highlight text={group.profileName} />
            </span>
            {group.isBaseline ? (
              <Badge appearance="tint" color="success">
                Baseline
              </Badge>
            ) : (
              <Badge appearance="tint" color="brand">
                Security Profile
              </Badge>
            )}
          </div>
          <span className={styles.counts}>
            {group.policyCount} {group.policyCount === 1 ? 'policy' : 'policies'} ·{' '}
            {ruleCount} {ruleCount === 1 ? 'rule' : 'rules'}
          </span>
        </div>

        <div className={styles.actionLane}>
          <ProfileSummaryChips group={group} />
        </div>

        <div className={styles.caLane}>
          <ProfileTargets group={group} />
        </div>
      </div>

      {expanded && <ProfileBody rows={rows} forceExpand={forceExpandPolicies} />}
    </section>
  );
}

// ---------------------------------------------------------------------------
// Details drawer (double-click → "Policy details")
// ---------------------------------------------------------------------------

/** A resolved object ready to render in the details drawer. */
interface DetailView {
  kindLabel: string;
  title: string;
  summary: { label: string; value: string }[];
  raw: unknown;
}

function fmt(value: unknown): string {
  if (value === null || value === undefined || value === '') return '—';
  if (Array.isArray(value)) return value.length === 0 ? '—' : String(value.length);
  return String(value);
}

function buildProfileDetail(profile: SecurityProfile): DetailView {
  const links = profile.policies ?? [];
  const ruleCount = links.reduce(
    (n, l) => n + ((l.policy as { policyRules?: unknown[] } | undefined)?.policyRules?.length ?? 0),
    0,
  );
  const ca = (profile.conditionalAccessPolicies ?? [])
    .map((c) => c.displayName ?? c.id)
    .join(', ');
  return {
    kindLabel: 'Security Profile',
    title: profile.name ?? '(unnamed profile)',
    summary: [
      { label: 'Priority', value: fmt(profile.priority) },
      { label: 'State', value: fmt(profile.state) },
      { label: 'Description', value: fmt(profile.description) },
      { label: 'Policies', value: String(links.length) },
      { label: 'Rules', value: String(ruleCount) },
      { label: 'Conditional Access', value: ca || '—' },
      { label: 'Created', value: fmt(profile.createdDateTime) },
      { label: 'Last modified', value: fmt(profile.lastModifiedDateTime) },
      { label: 'ID', value: fmt(profile.id) },
    ],
    raw: profile,
  };
}

function buildPolicyDetail(link: FilteringPolicyLink): DetailView {
  const policy = link.policy as
    | { name?: string; action?: string; version?: string; policyRules?: unknown[]; settings?: { defaultAction?: string } }
    | undefined;
  const typeLabel =
    resolvePolicyTypeLabel(link['@odata.type'], (policy as { '@odata.type'?: string } | undefined)?.['@odata.type']) ??
    'Unsupported type';
  const action = policy?.action ?? policy?.settings?.defaultAction ?? link.action;
  return {
    kindLabel: typeLabel,
    title: policy?.name ?? '(unnamed policy)',
    summary: [
      { label: 'Type', value: typeLabel },
      { label: 'Action / default', value: fmt(action) },
      { label: 'Link state', value: fmt(link.state) },
      { label: 'Priority', value: fmt(link.priority) },
      { label: 'Version', value: fmt(policy?.version) },
      { label: 'Rules', value: String(policy?.policyRules?.length ?? 0) },
    ],
    raw: link,
  };
}

function buildRuleDetail(rule: Record<string, unknown>, typeLabel: string): DetailView {
  const settings = rule.settings as { status?: string } | undefined;
  return {
    kindLabel: `${typeLabel} rule`,
    title: (rule.name as string) ?? '(unnamed rule)',
    summary: [
      { label: 'Priority', value: fmt(rule.priority) },
      { label: 'Action', value: fmt(rule.action) },
      { label: 'Status', value: fmt(settings?.status) },
      { label: 'Rule type', value: fmt(rule.ruleType) },
      { label: 'ID', value: fmt(rule.id) },
    ],
    raw: rule,
  };
}

function PolicyDetailsDrawer({
  view,
  onClose,
}: {
  view: DetailView | null;
  onClose: () => void;
}) {
  const styles = useStyles();
  return (
    <OverlayDrawer
      position="end"
      open={view !== null}
      onOpenChange={(_, d) => {
        if (!d.open) onClose();
      }}
      size="medium"
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close"
              icon={<Dismiss24Regular />}
              onClick={onClose}
            />
          }
        >
          {view && (
            <span className={styles.detailHeaderTitle}>
              <Badge appearance="tint" color="brand">
                {view.kindLabel}
              </Badge>
              <span>{view.title}</span>
            </span>
          )}
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody>
        {view && (
          <div className={styles.detailBody}>
            <div className={styles.detailGrid}>
              {view.summary.map((s) => (
                <div key={s.label} className={styles.detailRow}>
                  <span className={styles.detailKey}>{s.label}</span>
                  <span className={styles.detailValue}>{s.value}</span>
                </div>
              ))}
            </div>
            <Text className={styles.detailJsonLabel}>Raw JSON (all details)</Text>
            <pre className={styles.detailJson}>
              {JSON.stringify(view.raw, null, 2)}
            </pre>
          </div>
        )}
      </DrawerBody>
    </OverlayDrawer>
  );
}

// ---------------------------------------------------------------------------
// Public component
// ---------------------------------------------------------------------------

export function PolicyTable({
  groups,
  profiles,
}: {
  groups: ProfileGroup[];
  profiles: SecurityProfile[];
}) {
  const styles = useStyles();
  const [filters, setFilters] = useState<TableFilters>(EMPTY_FILTERS);
  // Start with all profiles collapsed so the user sees the full inventory at a glance.
  const [collapsed, setCollapsed] = useState<Set<string>>(
    () => new Set(groups.map((g) => g.profileId)),
  );
  const [detail, setDetail] = useState<DetailView | null>(null);

  const profileById = useMemo(
    () => new Map(profiles.map((p) => [p.id, p])),
    [profiles],
  );

  const detailOpeners = useMemo<DetailOpeners>(
    () => ({
      openProfile: (profileId) => {
        const p = profileById.get(profileId);
        if (p) setDetail(buildProfileDetail(p));
      },
      openPolicy: (policyKey) => {
        const [profileId, linkId] = policyKey.split('/');
        const p = profileById.get(profileId);
        const link = (p?.policies ?? []).find((l) => l.id === linkId);
        if (link) setDetail(buildPolicyDetail(link));
      },
      openRule: (rowKey) => {
        const [profileId, linkId, ruleId] = rowKey.split('/');
        const p = profileById.get(profileId);
        const link = (p?.policies ?? []).find((l) => l.id === linkId);
        const rule = (
          (link?.policy as { policyRules?: Record<string, unknown>[] } | undefined)
            ?.policyRules ?? []
        ).find((r) => r.id === ruleId);
        if (link && rule) {
          const typeLabel =
            resolvePolicyTypeLabel(
              link['@odata.type'],
              (link.policy as { '@odata.type'?: string } | undefined)?.['@odata.type'],
            ) ?? 'Policy';
          setDetail(buildRuleDetail(rule, typeLabel));
        }
      },
    }),
    [profileById],
  );

  const data = useMemo(() => groups.flatMap((g) => g.rows), [groups]);
  const groupById = useMemo(
    () => new Map(groups.map((g) => [g.profileId, g])),
    [groups],
  );

  // Distinct Action / Type values present in the data, for the filter dropdowns.
  const actionOptions = useMemo(
    () => [...new Set(data.map((r) => r.action).filter(Boolean))].sort(),
    [data],
  );
  const typeOptions = useMemo(
    () => [...new Set(data.map((r) => r.ruleType).filter(Boolean))].sort(),
    [data],
  );

  const columns = useMemo<ColumnDef<PolicyRow>[]>(
    () => [
      { accessorKey: 'profilePriority' },
      { accessorKey: 'policyPriority' },
    ],
    [],
  );

  const sorting: SortingState = useMemo(
    () => [
      { id: 'profilePriority', desc: false },
      { id: 'policyPriority', desc: false },
    ],
    [],
  );

  const globalFilterFn = useMemo(() => makeGlobalFilter(groupById), [groupById]);

  const table = useReactTable({
    data,
    columns,
    state: { sorting, globalFilter: filters },
    onGlobalFilterChange: (updater) =>
      setFilters((prev) =>
        typeof updater === 'function'
          ? (updater as (old: TableFilters) => TableFilters)(prev)
          : (updater as TableFilters),
      ),
    globalFilterFn,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
  });

  // Group the filtered/sorted leaf rows back by profile, preserving order.
  const sections = useMemo(() => {
    const out: { group: ProfileGroup; rows: PolicyRow[] }[] = [];
    const indexByProfile = new Map<string, number>();
    for (const r of table.getRowModel().rows) {
      const row = r.original;
      const group = groupById.get(row.profileId);
      if (!group) continue;
      let idx = indexByProfile.get(row.profileId);
      if (idx === undefined) {
        idx = out.length;
        indexByProfile.set(row.profileId, idx);
        out.push({ group, rows: [] });
      }
      out[idx].rows.push(row);
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table, groupById, filters]);

  const filtering = filtersActive(filters);
  const totalRules = data.filter((r) => r.ruleName !== '').length;
  const visibleRules = sections.reduce(
    (n, s) => n + s.rows.filter((r) => r.ruleName !== '').length,
    0,
  );

  const isExpanded = (profileId: string) =>
    filtering || !collapsed.has(profileId);

  const toggle = (profileId: string) =>
    setCollapsed((prev) => {
      const next = new Set(prev);
      if (next.has(profileId)) next.delete(profileId);
      else next.add(profileId);
      return next;
    });

  const expandAll = () => setCollapsed(new Set());
  const collapseAll = () => setCollapsed(new Set(groups.map((g) => g.profileId)));

  const setField = (field: keyof TableFilters, value: string) =>
    setFilters((prev) => ({ ...prev, [field]: value }));

  return (
    <DetailContext.Provider value={detailOpeners}>
    <div className={styles.root}>
      <div className={styles.toolbar}>
        <Input
          className={styles.search}
          contentBefore={<Search16Regular />}
          placeholder="Search profiles, policies, rules, destinations, CA…"
          value={filters.search}
          onChange={(_, d) => setField('search', d.value)}
        />
        <Dropdown
          className={styles.filterDropdown}
          placeholder="All actions"
          value={filters.action || 'All actions'}
          selectedOptions={filters.action ? [filters.action] : ['__all']}
          onOptionSelect={(_, d) =>
            setField('action', d.optionValue === '__all' ? '' : d.optionValue ?? '')
          }
        >
          <Option value="__all">All actions</Option>
          {actionOptions.map((a) => (
            <Option key={a} value={a}>
              {a}
            </Option>
          ))}
        </Dropdown>
        <Dropdown
          className={styles.filterDropdown}
          placeholder="All types"
          value={filters.ruleType || 'All types'}
          selectedOptions={filters.ruleType ? [filters.ruleType] : ['__all']}
          onOptionSelect={(_, d) =>
            setField('ruleType', d.optionValue === '__all' ? '' : d.optionValue ?? '')
          }
        >
          <Option value="__all">All types</Option>
          {typeOptions.map((t) => (
            <Option key={t} value={t}>
              {t}
            </Option>
          ))}
        </Dropdown>
        <Input
          className={styles.ruleFilter}
          placeholder="Filter rule name…"
          value={filters.rule}
          onChange={(_, d) => setField('rule', d.value)}
        />
        <Text size={200} className={styles.count}>
          {sections.length} of {groups.length} profiles · {visibleRules} of{' '}
          {totalRules} rules
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
        <Button
          size="small"
          appearance="subtle"
          disabled={filtering}
          onClick={collapseAll}
        >
          Collapse all
        </Button>
      </div>

      <div className={styles.scroll}>
        {sections.length === 0 ? (
          <div className={styles.noMatches}>
            No profiles match the current filter.
          </div>
        ) : (
          <SearchHighlightContext.Provider value={filters.search}>
            {sections.map(({ group, rows }) => (
              <ProfileSection
                key={group.profileId}
                group={group}
                rows={rows}
                expanded={isExpanded(group.profileId)}
                forceExpandPolicies={filtering}
                onToggle={() => toggle(group.profileId)}
              />
            ))}
          </SearchHighlightContext.Provider>
        )}
      </div>
    </div>
    <PolicyDetailsDrawer view={detail} onClose={() => setDetail(null)} />
    </DetailContext.Provider>
  );
}
