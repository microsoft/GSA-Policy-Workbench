/**
 * What-If panel — Tier 1 (UI) — the effective-policy resolver.
 *
 * Lets an admin ask "which Internet Access policy applies to user X and/or
 * destination Y?". It is a pure display view: it reads the already-loaded
 * profile tree + CA detail and calls the pure `resolveEffective` resolver.
 * No Graph calls, no scopes, no Tier 2/3 changes — display only.
 */

import { useEffect, useMemo, useState } from 'react';
import {
  Badge,
  Button,
  Dropdown,
  Input,
  Option,
  Spinner,
  Text,
  Tooltip,
  makeStyles,
  mergeClasses,
  tokens,
  type BadgeProps,
} from '@fluentui/react-components';
import {
  CheckmarkCircle20Filled,
  DismissCircle20Filled,
  QuestionCircle20Regular,
  Person20Regular,
  Globe20Regular,
  Target20Regular,
  LockClosed16Regular,
  Server16Regular,
  Tag20Regular,
  Tag16Regular,
  ChevronDown16Regular,
  ChevronRight16Regular,
  Warning20Filled,
  ShieldCheckmark16Regular,
  TextBulletListSquare20Regular,
} from '@fluentui/react-icons';
import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type { ConditionalAccessPolicy, DirectoryObjectRef } from '../../model/definitions/ConditionalAccessPolicy.definition';
import type { WebCategory } from '../../model/definitions/WebContentFilteringPolicy.definition';
import type { ForwardingProfile } from '../../model/definitions/ForwardingProfile.definition';
import type { PrivateAccessDomain } from '../../adapters/graph/loader';
import { useDataSource } from '../../app/dataSourceContext';
import { useWebCategoryLookup } from '../../query/hooks/useWebCategoryLookup';
import { useResolveUser } from '../../query/hooks/useResolveUser';
import { debugGroup, debugLog, debugTable, isDebugEnabled } from '../../app/debugLog';
import { buildPrivateAccessViewModel } from '../private/privateAccessRows';
import {
  resolvePrivateAccessDestination,
  type PrivateAccessAppMatch,
  type SegmentHit,
} from './privateAccessWhatIf';
import {
  resolveEffective,
  type AcquisitionEvaluation,
  type Applicability,
  type PolicyAction,
  type ProfileEvaluation,
} from './effectivePolicy';
import { actionColor } from '../actionColor';
import { friendlyError } from '../friendlyError';
import { deriveWhatIfTrace } from './whatifTrace';
import { WhatIfTraceDrawer } from './WhatIfTraceDrawer';
import { tracer } from '../../app/tracer';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
  },
  inputs: {
    display: 'flex',
    gap: tokens.spacingHorizontalL,
    flexWrap: 'wrap',
    alignItems: 'flex-end',
  },
  field: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    minWidth: '260px',
  },
  fieldWide: {
    flexGrow: 1,
    minWidth: '540px',
  },
  fieldLabel: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
  },
  hint: { color: tokens.colorNeutralForeground3 },

  // Resolved-user info box (live mode, shown between inputs row and outcome) -
  userStatus: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    flexWrap: 'wrap' as const,
  },

  // Web content category lookup (Internet Access, live-only) -----------------
  categoryBlock: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalL,
    borderRadius: tokens.borderRadiusXLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  categoryHead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    cursor: 'pointer',
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
  },
  categoryHeadGrow: { flex: 1 },
  categoryChevron: { color: tokens.colorNeutralForeground2, display: 'inline-flex' },
  categoryRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    alignItems: 'center',
    flexWrap: 'wrap',
  },
  categoryInput: { flexGrow: 1, minWidth: '320px' },
  categoryResult: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'center',
    flexWrap: 'wrap',
  },

  // Outcome banner -----------------------------------------------------------
  outcome: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  outcomeAllow: {
    background: `linear-gradient(90deg, ${tokens.colorPaletteGreenBackground2} 0%, ${tokens.colorNeutralBackground1} 80%)`,
    borderLeft: `5px solid ${tokens.colorPaletteGreenBorderActive}`,
  },
  outcomeBlock: {
    background: `linear-gradient(90deg, ${tokens.colorPaletteRedBackground2} 0%, ${tokens.colorNeutralBackground1} 80%)`,
    borderLeft: `5px solid ${tokens.colorPaletteRedBorderActive}`,
  },
  outcomeNeutral: {
    background: tokens.colorNeutralBackground2,
    borderLeft: `5px solid ${tokens.colorNeutralStroke1}`,
  },
  outcomeEscape: {
    background: `linear-gradient(90deg, ${tokens.colorPaletteDarkOrangeBackground2} 0%, ${tokens.colorNeutralBackground1} 80%)`,
    borderLeft: `5px solid ${tokens.colorPaletteDarkOrangeBorderActive}`,
  },
  outcomeIcon: { display: 'flex', flexShrink: 0 },
  outcomeBody: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS },
  outcomeVerdict: { fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightBold },
  outcomeWhy: { color: tokens.colorNeutralForeground2 },
  allowText: { color: tokens.colorPaletteGreenForeground1 },
  blockText: { color: tokens.colorPaletteRedForeground1 },

  // Applicable profiles list -------------------------------------------------
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
  },
  sectionToggle: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    cursor: 'pointer',
    userSelect: 'none',
    color: tokens.colorNeutralForeground2,
  },
  sectionChevron: { color: tokens.colorNeutralForeground2, display: 'inline-flex' },
  profileList: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  profileRow: {
    display: 'grid',
    gridTemplateColumns: 'auto 1fr auto',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  profileRowWinner: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
    boxShadow: tokens.shadow4,
  },
  profileRowDimmed: { opacity: 0.55 },
  profileName: { fontWeight: tokens.fontWeightSemibold },
  reason: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  matchLine: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
  },
});

const APPLICABILITY_META: Record<
  Applicability,
  { color: BadgeProps['color']; label: string }
> = {
  yes: { color: 'success', label: 'Applies' },
  maybe: { color: 'warning', label: 'May apply' },
  unknown: { color: 'subtle', label: 'Unknown' },
  no: { color: 'danger', label: "Doesn't apply" },
};

function actionBadge(action: PolicyAction) {
  if (action === '') return null;
  return (
    <Badge appearance="filled" color={actionColor(action)}>
      {action}
    </Badge>
  );
}

function ApplicabilityIcon({ value }: { value: Applicability }) {
  if (value === 'yes') return <CheckmarkCircle20Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />;
  if (value === 'no') return <DismissCircle20Filled style={{ color: tokens.colorPaletteRedForeground1 }} />;
  return <QuestionCircle20Regular style={{ color: tokens.colorNeutralForeground3 }} />;
}

interface WhatIfPanelProps {
  profiles: SecurityProfile[];
  caDetails: ConditionalAccessPolicy[];
  directory: DirectoryObjectRef[];
  privateAccess?: PrivateAccessDomain;
  /**
   * GSA traffic-forwarding profiles (state, priority, acquisition rules) —
   * the client-side "stage 1" of What-If. An
   * empty array degrades acquisition verdicts to `unknown` (no false escape
   * flags when this data isn't captured by the active source).
   */
  forwardingProfiles?: ForwardingProfile[];
  /**
   * Whether to offer the Internet Access web-category lookup. Hidden in the
   * Private Access view (it is an IA-only concept).
   */
  showWebCategory?: boolean;
}

export function WhatIfPanel({
  profiles,
  caDetails,
  directory,
  privateAccess,
  forwardingProfiles = [],
  showWebCategory = true,
}: WhatIfPanelProps) {
  const styles = useStyles();
  const { mode } = useDataSource();
  const isLive = mode === 'graph';
  const resolveUser = useResolveUser();
  const resolveCategory = useWebCategoryLookup(); // auto-fires on Evaluate in live mode
  const [userId, setUserId] = useState('');
  const [upnInput, setUpnInput] = useState('');          // live-mode UPN field
  const [destinationInput, setDestinationInput] = useState(''); // live field value
  const [destination, setDestination] = useState('');           // last evaluated value
  const [profilesOpen, setProfilesOpen] = useState(false);
  const [traceOpen, setTraceOpen] = useState(false);

  const users = useMemo(
    () =>
      directory
        .filter((o) => (o['@odata.type'] ?? '').includes('user'))
        .sort((a, b) =>
          (a.displayName ?? a.id).localeCompare(b.displayName ?? b.id),
        ),
    [directory],
  );

  const userLabel = useMemo(
    () => users.find((u) => u.id === userId)?.displayName ?? undefined,
    [users, userId],
  );

  // In live mode the resolver uses the Graph-resolved user; in file mode the dropdown.
  const effectiveUserId    = isLive ? (resolveUser.data?.id ?? '') : userId;
  const effectiveUserLabel = isLive ? resolveUser.data?.displayName : userLabel;
  const effectiveGroupIds  = isLive ? resolveUser.data?.groupIds : undefined;

  const effectiveCategoryName = isLive
    ? (resolveCategory.data?.displayName ?? resolveCategory.data?.name)
    : undefined;

  const result = useMemo(
    () => resolveEffective(
      profiles, caDetails,
      { userId: effectiveUserId, userLabel: effectiveUserLabel, destination, groupIds: effectiveGroupIds, resolvedCategoryName: effectiveCategoryName },
      forwardingProfiles,
    ),
    [profiles, caDetails, effectiveUserId, effectiveUserLabel, destination, effectiveGroupIds, effectiveCategoryName, forwardingProfiles],
  );

  const paResult = useMemo(() => {
    if (!privateAccess) return { hasDestination: false, matches: [] };
    const vm = buildPrivateAccessViewModel(
      privateAccess.apps,
      privateAccess.appProxyApps,
      privateAccess.authStrength,
      caDetails,
    );
    return resolvePrivateAccessDestination(vm.apps, destination, forwardingProfiles);
  }, [privateAccess, caDetails, destination, forwardingProfiles]);

  const paAppCount = privateAccess?.apps.length ?? 0;

  const active = result.hasUser || result.hasDestination;

  const traceRecords = useMemo(
    () => active ? deriveWhatIfTrace(
      result,
      paAppCount > 0 ? paResult : undefined,
      {
        resolvedCategoryName: effectiveCategoryName,
        resolvedCategoryGroup: resolveCategory.data?.group,
        userGroups: resolveUser.data?.groups,
      },
    ) : [],
    [active, result, paResult, paAppCount, effectiveCategoryName, resolveCategory.data, resolveUser.data],
  );

  // Profiles that apply to the chosen user (or all, if no user filter).
  const applicable = result.profiles.filter(
    (p) => p.applicability !== 'no',
  );

  // Write trace records to the tracer (console + buffer) on each resolution.
  useEffect(() => {
    if (!active) return;
    tracer.clear('whatif');
    tracer.trace(traceRecords);
  }, [active, traceRecords]);

  // Legacy F12 debug output (window.__gsaDebug.on()).
  useEffect(() => {
    if (!active || !isDebugEnabled()) return;
    const outcome = result.winner
      ? `${result.winner.policy.action || 'match'}`
      : result.hasDestination
        ? 'no rule match (default)'
        : '(profiles only)';
    debugGroup(
      `[GSA] What-If · user=${(effectiveUserLabel ?? effectiveUserId) || '(any)'} dest=${destination || '(none)'} → ${outcome}`,
      () => {
        debugLog('inputs', { userId: effectiveUserId, userLabel: effectiveUserLabel, destination, groupIds: effectiveGroupIds?.length });
        if (result.acquisition) {
          debugLog('acquisition (internet)', result.acquisition);
        }
        if (result.winner) {
          debugLog('winner', {
            profile: result.winner.profile.profileName,
            profilePriority: result.winner.profile.priority,
            policy: result.winner.policy.policyName,
            action: result.winner.policy.action,
            matchedRule: result.winner.policy.match?.ruleName,
            matchedDestination: result.winner.policy.match?.destinationLabel,
          });
        }
        debugTable(
          result.profiles.map((p) => ({
            profile: p.profileName,
            priority: p.priority ?? '',
            applicability: p.applicability,
            reason: p.applicabilityReason ?? '',
            firstMatch: p.firstMatch?.match?.ruleName ?? '',
            winner: result.winner?.profile.profileId === p.profileId ? '★' : '',
          })),
        );
        if (paAppCount > 0) {
          debugLog('private-access matches', paResult.matches.length);
          if (paResult.acquisition) {
            debugLog('acquisition (private)', paResult.acquisition);
          }
          debugTable(
            paResult.matches.map((m) => ({
              app: m.app.name,
              segment: m.hits[0]?.segment.destinationHost ?? '',
              reason: m.hits[0]?.reason ?? '',
              access: m.access,
            })),
          );
        }
      },
    );
  }, [active, result, paResult, effectiveUserId, effectiveUserLabel, effectiveGroupIds, destination, paAppCount]);

  const evaluate = () => {
    const dest = destinationInput.trim();
    setDestination(dest);
    if (isLive) {
      const upn = upnInput.trim();
      if (upn) resolveUser.mutate(upn);
      else resolveUser.reset();
      if (dest) resolveCategory.mutate(dest);
      else resolveCategory.reset();
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.inputs}>
        {isLive ? (
          <div className={styles.field}>
            <Text className={styles.fieldLabel}>
              <Person20Regular /> User
            </Text>
            <Input
              placeholder="User Principal Name — e.g. user@contoso.com"
              value={upnInput}
              onChange={(_, d) => setUpnInput(d.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') evaluate(); }}
            />
          </div>
        ) : (
          <div className={styles.field}>
            <Text className={styles.fieldLabel}>
              <Person20Regular /> User
            </Text>
            <Dropdown
              placeholder={users.length ? 'Any user (no filter)' : 'No directory data in this source'}
              disabled={users.length === 0}
              value={userLabel ?? 'Any user (no filter)'}
              selectedOptions={userId ? [userId] : ['__any']}
              onOptionSelect={(_, d) => setUserId(d.optionValue === '__any' ? '' : d.optionValue ?? '')}
            >
              <Option value="__any">Any user (no filter)</Option>
              {users.map((u) => (
                <Option key={u.id} value={u.id} text={u.displayName ?? u.id}>
                  {u.displayName ?? u.id}
                  {u.userPrincipalName ? ` · ${u.userPrincipalName}` : ''}
                </Option>
              ))}
            </Dropdown>
          </div>
        )}

        <div className={mergeClasses(styles.field, styles.fieldWide)}>
          <Text className={styles.fieldLabel}>
            <Globe20Regular /> Destination
          </Text>
          <div style={{ display: 'flex', gap: tokens.spacingHorizontalS }}>
            <Input
              style={{ flexGrow: 1 }}
              placeholder="FQDN, IP, CIDR or wildcard — e.g. *.contoso.com, 192.168.110.0/24"
              value={destinationInput}
              onChange={(_, d) => setDestinationInput(d.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') evaluate(); }}
            />
            <Button
              appearance="primary"
              disabled={!destinationInput.trim()}
              onClick={evaluate}
            >
              Evaluate
            </Button>
          </div>
        </div>

        {(userId || upnInput || destinationInput || destination) && (
          <Button
            appearance="subtle"
            onClick={() => {
              setUserId('');
              setUpnInput('');
              setDestinationInput('');
              setDestination('');
              resolveUser.reset();
              resolveCategory.reset();
            }}
          >
            Clear
          </Button>
        )}
        <Tooltip content="Explain this decision" relationship="label">
          <Button
            appearance="subtle"
            icon={<TextBulletListSquare20Regular />}
            aria-label="Explain this decision"
            disabled={!active}
            onClick={() => setTraceOpen(true)}
          />
        </Tooltip>
      </div>

      {showWebCategory && <WebCategoryLookup />}

      {isLive && (resolveUser.isPending || resolveUser.isError || resolveUser.data) && (
        <div className={styles.userStatus}>
          {resolveUser.isPending && (
            <><Spinner size="tiny" /><Text className={styles.reason}>Resolving {upnInput}…</Text></>
          )}
          {resolveUser.isError && (
            <><DismissCircle20Filled style={{ color: tokens.colorPaletteRedForeground1, flexShrink: 0 }} />
            <Text style={{ color: tokens.colorPaletteRedForeground1 }}>
              {friendlyError(resolveUser.error, 'Could not resolve that user.')}
            </Text></>
          )}
          {resolveUser.data && !resolveUser.isPending && (
            <><Person20Regular style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
            <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{resolveUser.data.displayName}</Text>
            <Text className={styles.reason}>{resolveUser.data.userPrincipalName}</Text>
            {resolveUser.data.groupIds.length > 0 && (
              <Badge appearance="tint" color="informative">
                {resolveUser.data.groupIds.length} group{resolveUser.data.groupIds.length === 1 ? '' : 's'} resolved
              </Badge>
            )}</>
          )}
        </div>
      )}

      {isLive && destination && (resolveCategory.isPending || resolveCategory.isError || resolveCategory.data) && (
        <div className={styles.userStatus}>
          {resolveCategory.isPending && (
            <><Spinner size="tiny" /><Text className={styles.reason}>Resolving category…</Text></>
          )}
          {resolveCategory.isError && (
            <><Tag20Regular style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
            <Text className={styles.reason}>
              {friendlyError(resolveCategory.error, 'Category unavailable.')}
            </Text></>
          )}
          {resolveCategory.data && !resolveCategory.isPending && (
            <><Tag20Regular style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
            <Text style={{ fontWeight: tokens.fontWeightSemibold }}>
              {resolveCategory.data.displayName ?? resolveCategory.data.name}
            </Text>
            <Text className={styles.reason}>{destination}</Text>
            {resolveCategory.data.group && (
              <Badge appearance="tint" color="informative">{resolveCategory.data.group}</Badge>
            )}</>
          )}
        </div>
      )}

      {!active ? (
        <Text className={styles.hint}>
          Pick a user, type a destination and press Evaluate (or Enter), or both — the effective
          outcome appears here. Evaluation follows GSA order: profile priority, then policy
          priority, first match wins.
        </Text>
      ) : (
        <>
          {result.hasDestination && result.acquisition && (
            <AcquisitionBanner acquisition={result.acquisition} trafficLabel="Internet Access" />
          )}
          <OutcomeBanner result={result} resolvedCategoryName={effectiveCategoryName} />
          <div
            className={styles.sectionToggle}
            role="button"
            tabIndex={0}
            aria-expanded={profilesOpen}
            onClick={() => setProfilesOpen((v) => !v)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                setProfilesOpen((v) => !v);
              }
            }}
          >
            <span className={styles.sectionChevron}>
              {profilesOpen ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
            </span>
            <Text className={styles.sectionTitle}>
              {result.hasUser
                ? `Profiles that apply to ${userLabel ?? 'this user'} (${applicable.length})`
                : `Profiles evaluated (${applicable.length})`}
            </Text>
          </div>
          {profilesOpen && (
            <div className={styles.profileList}>
              {applicable.map((profile) => (
                <ProfileEvalRow
                  key={profile.profileId}
                  profile={profile}
                  isWinner={result.winner?.profile.profileId === profile.profileId}
                  showMatch={result.hasDestination}
                />
              ))}
              {applicable.length === 0 && (
                <Text className={styles.reason}>
                  No Security Profile applies to this user (no linked Conditional Access
                  policy targets them). Only the Baseline profile would apply.
                </Text>
              )}
            </div>
          )}

          {result.hasDestination && paAppCount > 0 && (
            <>
              {paResult.acquisition && (
                <AcquisitionBanner acquisition={paResult.acquisition} trafficLabel="Private Access" />
              )}
              <PrivateAccessMatches result={paResult} />
            </>
          )}
        </>
      )}
      <WhatIfTraceDrawer
        open={traceOpen}
        onClose={() => setTraceOpen(false)}
        records={traceRecords}
      />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Web content category lookup (Internet Access) — live-only
// ---------------------------------------------------------------------------

/**
 * Optional Internet Access helper: resolve which web content filtering category
 * a URL/host maps to, via the beta web category checker (preview). Live-only —
 * there is no fixture for arbitrary URL lookups, so it is disabled in file mode.
 */
export function WebCategoryLookup() {
  const styles = useStyles();
  const { mode } = useDataSource();
  const lookup = useWebCategoryLookup();
  const [url, setUrl] = useState('');
  const live = mode === 'graph';
  // Collapsed by default in file mode (where it can't run); open in live mode.
  const [open, setOpen] = useState(live);

  const submit = () => {
    const q = url.trim();
    if (q) lookup.mutate(q);
  };

  return (
    <div className={styles.categoryBlock}>
      <div
        className={styles.categoryHead}
        role="button"
        tabIndex={0}
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            setOpen((v) => !v);
          }
        }}
      >
        <Tag20Regular /> Web content category lookup
        <Badge appearance="outline" color="brand">
          Internet Access
        </Badge>
        {!live && (
          <Badge appearance="tint" color="subtle">
            Sign in to use
          </Badge>
        )}
        <span className={styles.categoryHeadGrow} />
        <span className={styles.categoryChevron}>
          {open ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
        </span>
      </div>
      {open && (
        <>
          <div className={styles.categoryRow}>
            <Input
              className={styles.categoryInput}
              placeholder="URL or host — e.g. msn.com/en-us/sports"
              value={url}
              disabled={!live}
              onChange={(_, d) => setUrl(d.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit();
              }}
            />
            <Button
              appearance="primary"
              disabled={!live || url.trim() === '' || lookup.isPending}
              onClick={submit}
            >
              {lookup.isPending ? <Spinner size="tiny" /> : 'Check category'}
            </Button>
          </div>
          {!live ? (
            <Text className={styles.hint}>
              Category lookup queries Microsoft Graph live — sign in to use it. It
              is not available in file mode.
            </Text>
          ) : lookup.isError ? (
            <Text className={styles.blockText}>
              {friendlyError(lookup.error, 'The category lookup failed.')}
            </Text>
          ) : lookup.data ? (
            <CategoryResult data={lookup.data} />
          ) : (
            <Text className={styles.hint}>
              Resolve which web content filtering category a URL or host maps to.
            </Text>
          )}
        </>
      )}
    </div>
  );
}

function CategoryResult({ data }: { data: WebCategory }) {
  const styles = useStyles();
  const label = data.displayName || data.name;
  if (!label) {
    return (
      <Text className={styles.hint}>No category was returned for that URL.</Text>
    );
  }
  return (
    <div className={styles.categoryResult}>
      <Badge appearance="filled" color="brand" icon={<Tag16Regular />}>
        {label}
      </Badge>
      {data.group && (
        <Badge appearance="tint" color="informative">
          {data.group}
        </Badge>
      )}
    </div>
  );
}

const PA_REASON_LABEL: Record<SegmentHit['reason'], string> = {
  fqdn: 'FQDN',
  'dns-suffix': 'DNS suffix',
  ip: 'IP',
  'ip-range': 'IP range',
  cidr: 'CIDR',
};

function paAccessBadge(access: PrivateAccessAppMatch['access']) {
  if (access === 'block') {
    return (
      <Badge appearance="filled" color="danger">
        block
      </Badge>
    );
  }
  if (access === 'grant') {
    return (
      <Badge appearance="filled" color="success">
        grant
      </Badge>
    );
  }
  return (
    <Badge appearance="outline" color="subtle">
      no CA
    </Badge>
  );
}

function PrivateAccessMatches({
  result,
}: {
  result: ReturnType<typeof resolvePrivateAccessDestination>;
}) {
  const styles = useStyles();
  const caveat = hypotheticalCaveat(result.acquisition, 'Private Access', styles);
  return (
    <>
      <Text className={styles.sectionTitle}>
        <LockClosed16Regular /> Private Access apps reachable ({result.matches.length})
      </Text>
      {result.matches.length === 0 ? (
        <Text className={styles.reason}>
          No Private Access application segment covers{' '}
          <strong>{result.destination}</strong>. Traffic to it would not enter the
          Private Access tunnel.
        </Text>
      ) : (
        <div className={styles.profileList}>
          {result.matches.map((m) => {
            const top = m.hits[0];
            return (
              <div key={m.app.id} className={styles.profileRow}>
                <Server16Regular style={{ color: tokens.colorBrandForeground1 }} />
                <div>
                  <div className={styles.profileName}>{m.app.name}</div>
                  <div className={mergeClasses(styles.matchLine, styles.reason)}>
                    <Badge appearance="tint" color="brand">
                      {PA_REASON_LABEL[top.reason]}
                    </Badge>
                    <span>
                      {top.segment.destinationHost}
                      {top.segment.ports !== '—' ? ` · ${top.segment.ports}` : ''}
                      {m.hits.length > 1 ? ` · +${m.hits.length - 1} more segment${m.hits.length - 1 === 1 ? '' : 's'}` : ''}
                    </span>
                  </div>
                  {m.supersededBy && m.supersededBy.length > 0 && (
                    <div className={styles.reason}>
                      ⚠ Not applied — {m.supersededBy.join(', ')} also covers this destination and
                      takes priority. GSA never routes this traffic through Quick Access.
                    </div>
                  )}
                </div>
                <div>{paAccessBadge(m.access)}</div>
              </div>
            );
          })}
        </div>
      )}
      {caveat && <div className={styles.reason}>{caveat}</div>}
    </>
  );
}

function OutcomeBanner({ result, resolvedCategoryName }: { result: ReturnType<typeof resolveEffective>; resolvedCategoryName?: string }) {
  const styles = useStyles();
  const escapeCaveat = hypotheticalCaveat(result.acquisition, 'Internet Access', styles);

  // Destination given + a winning rule found.
  if (result.hasDestination && result.winner) {
    const { profile, policy } = result.winner;
    const isBlock = policy.action === 'block';
    return (
      <div
        className={mergeClasses(
          styles.outcome,
          isBlock ? styles.outcomeBlock : styles.outcomeAllow,
        )}
      >
        <span className={styles.outcomeIcon}>
          {isBlock ? (
            <DismissCircle20Filled style={{ color: tokens.colorPaletteRedForeground1 }} />
          ) : (
            <CheckmarkCircle20Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />
          )}
        </span>
        <div className={styles.outcomeBody}>
          <span
            className={mergeClasses(
              styles.outcomeVerdict,
              isBlock ? styles.blockText : styles.allowText,
            )}
          >
            {isBlock ? 'Blocked' : 'Allowed'}
            {result.destination ? ` — ${result.destination}` : ''}
          </span>
          <span className={styles.outcomeWhy}>
            Matched <strong>{policy.match?.ruleName}</strong> in policy{' '}
            <strong>{policy.policyName}</strong> · profile{' '}
            <strong>{profile.profileName}</strong>
            {result.userLabel ? ` · for ${result.userLabel}` : ''}.
          </span>
          {resolvedCategoryName && (
            <span className={styles.outcomeWhy}>
              <Tag16Regular style={{ verticalAlign: 'middle', marginRight: tokens.spacingHorizontalXS }} />
              <strong>{result.destination}</strong>{' resolved to category '}<strong>{resolvedCategoryName}</strong>
            </span>
          )}
          {escapeCaveat}
        </div>
      </div>
    );
  }

  // Destination given but nothing matched in an applicable profile.
  if (result.hasDestination) {
    return (
      <div className={mergeClasses(styles.outcome, styles.outcomeNeutral)}>
        <span className={styles.outcomeIcon}>
          <Target20Regular style={{ color: tokens.colorNeutralForeground3 }} />
        </span>
        <div className={styles.outcomeBody}>
          <span className={styles.outcomeVerdict}>No explicit rule match</span>
          <span className={styles.outcomeWhy}>
            No filtering rule in an applicable profile matches{' '}
            <strong>{result.destination}</strong>. The Baseline catch-all default
            would decide the outcome.
          </span>
          {escapeCaveat}
        </div>
      </div>
    );
  }

  // User only — summarise applicability.
  const applies = result.profiles.filter((p) => p.applicability === 'yes');
  return (
    <div className={mergeClasses(styles.outcome, styles.outcomeNeutral)}>
      <span className={styles.outcomeIcon}>
        <Person20Regular style={{ color: tokens.colorNeutralForeground2 }} />
      </span>
      <div className={styles.outcomeBody}>
        <span className={styles.outcomeVerdict}>
          {applies.length} profile{applies.length === 1 ? '' : 's'} apply to{' '}
          {result.userLabel ?? 'this user'}
        </span>
        <span className={styles.outcomeWhy}>
          Add a destination to see the exact allow / block outcome.
        </span>
      </div>
    </div>
  );
}

/**
 * Client-side traffic-forwarding acquisition verdict (What-If "stage 1").
 * Rendered ABOVE the cloud-side outcome, since
 * whether traffic ever reaches Global Secure Access is decided before any
 * Conditional Access / filtering-rule evaluation applies.
 *
 * `bypassed` / `unmatched` are flagged prominently as a traffic **escape** —
 * the maintainer's explicit ask: an admin must never read a quiet screen and
 * assume GSA governs traffic that in fact never reached it. `forwarded` is a
 * quiet confirmation line (avoid banner fatigue for the common case).
 * `unknown` / `disabled`-without-match stay silent-but-honest: `unknown`
 * renders nothing (no acquisition data captured — never claim an escape we
 * can't support), `disabled` still counts as an escape (the whole traffic
 * class never reaches GSA).
 */
/**
 * Turn an `AcquisitionEvaluation` into the small "this outcome is
 * hypothetical" note rendered under a cloud-side outcome. Two distinct
 * reasons: a genuine traffic escape (isEscape) vs. a higher-precedence
 * workload pre-empting evaluation entirely (preemptedBy) — see
 * `evaluateAcquisition` in effectivePolicy.ts.
 */
function hypotheticalCaveat(
  acquisition: AcquisitionEvaluation | undefined,
  workloadLabel: string,
  styles: ReturnType<typeof useStyles>,
) {
  if (!acquisition) return null;
  if (acquisition.isEscape) {
    return (
      <span className={styles.outcomeWhy}>
        ⚠ Hypothetical: this assumes the traffic reaches Global Secure Access — see
        the traffic-escape warning above.
      </span>
    );
  }
  if (acquisition.preemptedBy) {
    const workload = acquisition.preemptedBy === 'entra' ? 'Microsoft Entra' : 'Microsoft 365';
    return (
      <span className={styles.outcomeWhy}>
        ℹ Hypothetical: this destination is actually handled by {workload} traffic
        forwarding, evaluated before {workloadLabel} — see the note above.
      </span>
    );
  }
  return null;
}

function AcquisitionBanner({
  acquisition,
  trafficLabel,
}: {
  acquisition: AcquisitionEvaluation;
  trafficLabel: string;
}) {
  const styles = useStyles();

  if (acquisition.verdict === 'unknown') return null;

  if (!acquisition.isEscape) {
    const catNote = acquisition.match?.category ? ` · ${acquisition.match.category}` : '';
    const groupNote = acquisition.match?.serviceGroup ? ` · ${acquisition.match.serviceGroup}` : '';
    return (
      <Text className={styles.reason}>
        <ShieldCheckmark16Regular /> {trafficLabel}: {acquisition.reason}{groupNote}{catNote}.
      </Text>
    );
  }

  return (
    <div className={mergeClasses(styles.outcome, styles.outcomeEscape)}>
      <span className={styles.outcomeIcon}>
        <Warning20Filled style={{ color: tokens.colorPaletteDarkOrangeForeground1 }} />
      </span>
      <div className={styles.outcomeBody}>
        <span
          className={styles.outcomeVerdict}
          style={{ color: tokens.colorPaletteDarkOrangeForeground1 }}
        >
          Traffic escape — not sent to Global Secure Access ({trafficLabel})
        </span>
        <span className={styles.outcomeWhy}>{acquisition.reason}.</span>
      </div>
    </div>
  );
}

function ProfileEvalRow({
  profile,
  isWinner,
  showMatch,
}: {
  profile: ProfileEvaluation;
  isWinner: boolean;
  showMatch: boolean;
}) {
  const styles = useStyles();
  const meta = APPLICABILITY_META[profile.applicability];
  const dimmed = showMatch && !profile.firstMatch && !isWinner;

  return (
    <div
      className={mergeClasses(
        styles.profileRow,
        isWinner && styles.profileRowWinner,
        dimmed && styles.profileRowDimmed,
      )}
    >
      <ApplicabilityIcon value={profile.applicability} />

      <div>
        <div className={styles.profileName}>
          {profile.profileName}
          {profile.isBaseline ? ' · Baseline' : ''}
          {profile.priority !== null ? ` · priority ${profile.priority}` : ''}
        </div>
        <div className={styles.reason}>
          <Tooltip content={profile.applicabilityReason} relationship="label">
            <span>{profile.applicabilityReason}</span>
          </Tooltip>
        </div>
        {showMatch && profile.firstMatch && (
          <div className={mergeClasses(styles.matchLine, styles.reason)}>
            {actionBadge(profile.firstMatch.action)}
            <span>
              {profile.firstMatch.match?.ruleName} → {profile.firstMatch.match?.destinationLabel}
            </span>
          </div>
        )}
      </div>

      <div>
        {isWinner ? (
          <Badge appearance="filled" color="brand">
            Effective
          </Badge>
        ) : (
          <Badge appearance="tint" color={meta.color}>
            {meta.label}
          </Badge>
        )}
      </div>
    </div>
  );
}
