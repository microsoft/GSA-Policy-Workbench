/**
 * Forwarding Profiles card — Tier 1 (UI), read-only.
 *
 * Renders one independent, collapsible card per forwarding profile.
 * Each card header shows: [Traffic type] · [Profile name] — no nesting,
 * no type-group headers. Profiles of the same type are separate cards.
 *
 * - `internet` profiles: state + clientFallbackAction badge (fail-close auto-expanded)
 * - `private`  profiles: state + app count from linkedApps[]
 * - `m365`     profiles: state only
 *
 * Degrades gracefully on GA tenants (1 profile per type, no preview fields).
 *
 * See docs/Support-Multiple-Forwarding-Profiles.md for the design.
 */

import { useState } from 'react';
import {
  Badge,
  Text,
  ToggleButton,
  makeStyles,
  mergeClasses,
  tokens,
  type BadgeProps,
} from '@fluentui/react-components';
import {
  CheckmarkCircle16Filled,
  Warning16Filled,
  Info16Regular,
  LockClosed16Regular,
  Globe16Regular,
  LockShield16Regular,
  ChevronDown16Regular,
  ChevronRight16Regular,
  Apps16Regular,
  Link16Regular,
  Server16Regular,
} from '@fluentui/react-icons';
import type {
  ForwardingProfile,
  ForwardingRule,
  TrafficForwardingType,
} from '../../model/definitions/ForwardingProfile.definition';
import type { RuleDestination } from '../../model/definitions/FilteringRule.definition';
import type { PrivateAccessApp } from '../../model/definitions/PrivateAccessApp.definition';
import { appsInForwardingProfile } from '../private/privateAccessRows';

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  list: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  card: {
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: 'hidden',
  },
  cardFailClose: {
    borderTopColor: tokens.colorPaletteMarigoldBorder2,
    borderRightColor: tokens.colorPaletteMarigoldBorder2,
    borderBottomColor: tokens.colorPaletteMarigoldBorder2,
    borderLeftColor: tokens.colorPaletteMarigoldBorder2,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    cursor: 'pointer',
    userSelect: 'none',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  typeChunk: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexShrink: 0,
  },
  typeIcon: {
    color: tokens.colorNeutralForeground3,
    display: 'inline-flex',
  },
  typeLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
  },
  separator: {
    color: tokens.colorNeutralForeground4,
    fontSize: tokens.fontSizeBase200,
    flexShrink: 0,
  },
  profileName: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    flex: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  badges: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexShrink: 0,
  },
  priority: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  chevron: {
    color: tokens.colorNeutralForeground3,
    display: 'inline-flex',
    flexShrink: 0,
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalM,
    paddingLeft: tokens.spacingHorizontalXL,
    paddingRight: tokens.spacingHorizontalM,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  row: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
  },
  rowLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    minWidth: '120px',
    flexShrink: 0,
    paddingTop: '1px',
  },
  rowValue: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
  },
  rowValueWarn: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteMarigoldForeground2,
  },
  fallbackValue: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start',
    gap: tokens.spacingVerticalXXS,
  },
  appList: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: tokens.spacingHorizontalXS,
    rowGap: tokens.spacingVerticalXS,
  },
  serviceGroups: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  // sg* — security-profile-card visual language for service-group / named-policy rows
  sgCard: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusXLarge,
    overflow: 'hidden',
    boxShadow: tokens.shadow4,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  sgHeader: {
    display: 'grid',
    gridTemplateColumns: '48px minmax(0, 1fr) 160px',
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
  sgHeaderDisabled: {
    borderLeftColor: tokens.colorNeutralStroke1,
    background: `linear-gradient(90deg, ${tokens.colorNeutralBackground3} 0%, ${tokens.colorNeutralBackground1} 70%)`,
    opacity: 0.65,
    ':hover': {
      background: `linear-gradient(90deg, ${tokens.colorNeutralBackground3} 0%, ${tokens.colorNeutralBackground1Hover} 70%)`,
    },
  },
  sgGutter: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start',
    justifyContent: 'center',
    lineHeight: tokens.lineHeightBase200,
    minWidth: 0,
  },
  sgGutterLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
  },
  sgGutterValue: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  sgIdentity: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    minWidth: 0,
  },
  sgIdentityLead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  sgName: {
    fontWeight: tokens.fontWeightBold,
    fontSize: tokens.fontSizeBase400,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    minWidth: 0,
  },
  sgCounts: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: 'nowrap',
    paddingLeft: '24px',
  },
  sgAction: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-end',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
  },
  serviceGroupRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    padding: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalXS}`,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  serviceGroupRowDisabled: {
    opacity: 0.6,
  },
  serviceGroupName: {
    flex: 1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  serviceGroupCount: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  groupChevron: {
    color: tokens.colorNeutralForeground3,
    display: 'inline-flex',
    flexShrink: 0,
    marginLeft: 'auto',
  },
  ruleList: {
    display: 'flex',
    flexDirection: 'column',
    maxHeight: '320px',
    overflowY: 'auto',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: `0 0 ${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium}`,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderTop: 'none',
    marginTop: '-1px',
  },
  ruleGridRow: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) 116px 148px',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
  },
  ruleColTitles: {
    color: tokens.colorNeutralForeground3,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.03em',
    backgroundColor: tokens.colorNeutralBackground3,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    position: 'sticky',
    top: 0,
    zIndex: 1,
  },
  ruleRow: {
    borderTop: `1px solid ${tokens.colorNeutralStroke3}`,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  ruleRowAlt: { backgroundColor: tokens.colorNeutralBackground1 },
  ruleCell: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
  },
  ruleDest: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
});

// ---------------------------------------------------------------------------
// Static config
// ---------------------------------------------------------------------------

const TYPE_CONFIG: Record<string, { label: string; Icon: typeof Globe16Regular }> = {
  internet: { label: 'Internet Access', Icon: Globe16Regular },
  m365:     { label: 'Microsoft 365',   Icon: LockShield16Regular },
  private:  { label: 'Private Access',  Icon: LockClosed16Regular },
};

function stateBadge(state: string | undefined): {
  color: BadgeProps['color'];
  Icon: typeof CheckmarkCircle16Filled;
  label: string;
} {
  if (state === 'enabled')  return { color: 'success', Icon: CheckmarkCircle16Filled, label: 'on' };
  if (state === 'disabled') return { color: 'warning', Icon: Warning16Filled,         label: 'disabled' };
  return                           { color: 'subtle',  Icon: Info16Regular,            label: 'not configured' };
}

function destLabel(d: RuleDestination): string {
  if (d['@odata.type'] === '#microsoft.graph.networkaccess.ipRange') {
    const r = d as { beginAddress: string; endAddress: string };
    return `${r.beginAddress} – ${r.endAddress}`;
  }
  if ('value' in d && typeof (d as { value?: unknown }).value === 'string') {
    return (d as { value: string }).value;
  }
  return d['@odata.type']?.split('.').pop() ?? '(unknown)';
}

function rulePrimaryDest(rule: ForwardingRule): string {
  const dests = rule.destinations ?? [];
  if (dests.length === 0) return rule.name ?? rule.id;
  const first = destLabel(dests[0]);
  return dests.length > 1 ? `${first} +${dests.length - 1}` : first;
}

const FWDRULE_TYPE_META: Record<string, { Icon: typeof Globe16Regular; color: BadgeProps['color'] }> = {
  fqdn:      { Icon: Globe16Regular,  color: 'informative' },
  url:       { Icon: Link16Regular,   color: 'success'     },
  ipAddress: { Icon: Server16Regular, color: 'warning'     },
  ipSubnet:  { Icon: Server16Regular, color: 'warning'     },
  ipRange:   { Icon: Server16Regular, color: 'warning'     },
};

// ---------------------------------------------------------------------------
// Single profile card
// ---------------------------------------------------------------------------

interface ProfileCardProps {
  profile: ForwardingProfile;
  paApps: PrivateAccessApp[];
  showTypeLabel: boolean;
}

function ProfileCard({ profile, paApps, showTypeLabel }: ProfileCardProps) {
  const styles = useStyles();

  const type        = profile.trafficForwardingType ?? '';
  const isPrivate   = type === 'private';
  const isFailClose = profile.clientFallbackAction === 'block';

  // Private profiles start expanded so apps are immediately visible.
  // Fail-close profiles start expanded so the warning is immediately visible.
  // M365 profiles start expanded to show service-group state.
  const [open, setOpen] = useState(isPrivate || isFailClose || type === 'm365');
  const [expandedGroups, setExpandedGroups] = useState<ReadonlySet<string>>(new Set());

  const { color, Icon: StateIcon, label: stateLabel } = stateBadge(profile.state);
  const config   = TYPE_CONFIG[type] ?? { label: type, Icon: Globe16Regular };
  const TypeIcon = config.Icon;
  const Chevron  = open ? ChevronDown16Regular : ChevronRight16Regular;

  const inProfileApps = isPrivate ? appsInForwardingProfile(profile, paApps) : [];
  const excludedApps  = isPrivate && profile.linkedApps != null
    ? paApps.filter((a) => !inProfileApps.some((l) => l.appId === a.appId))
    : [];

  const assocCount    = profile.associations?.length ?? 0;
  const assignedLabel = assocCount === 0
    ? 'All users and devices'
    : `${assocCount} assignment(s) configured`;

  return (
    <div className={mergeClasses(styles.card, isFailClose && styles.cardFailClose)}>

      <div
        className={styles.header}
        onClick={() => setOpen((v) => !v)}
        role="button"
        aria-expanded={open}
      >
        {showTypeLabel && (
          <>
            <div className={styles.typeChunk}>
              <span className={styles.typeIcon}><TypeIcon /></span>
              <Text className={styles.typeLabel}>{config.label}</Text>
            </div>
            <Text className={styles.separator}>·</Text>
          </>
        )}

        <Text className={styles.profileName}>{profile.name ?? profile.id}</Text>

        <div className={styles.badges}>
          <Badge appearance="tint" color={color} icon={<StateIcon />}>
            {stateLabel}
          </Badge>

          {isFailClose && (
            <Badge appearance="tint" color="warning" icon={<LockClosed16Regular />}>
              fail-close
            </Badge>
          )}

          {isPrivate && profile.linkedApps && profile.linkedApps.length > 0 && (
            <Badge appearance="tint" color="informative" icon={<Apps16Regular />}>
              {inProfileApps.length} app{inProfileApps.length !== 1 ? 's' : ''}
            </Badge>
          )}
        </div>

        {profile.priority !== undefined && (
          <Text className={styles.priority}>Priority {profile.priority}</Text>
        )}

        <span className={styles.chevron}><Chevron /></span>
      </div>

      {open && (
        <div className={styles.body}>
          <div className={styles.row}>
            <Text className={styles.rowLabel}>Assigned to</Text>
            <Text className={styles.rowValue}>{assignedLabel}</Text>
          </div>

          <div className={styles.row}>
            <Text className={styles.rowLabel}>Client fallback</Text>
            <div className={styles.fallbackValue}>
              {/* Read-only toggle — pointerEvents:none keeps full button styling without
                  the disabled graying. Remove the wrapper and add onClick in V1+. */}
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
              {isFailClose ? (
                <Text className={styles.rowValueWarn}>
                  Traffic is dropped if the GSA edge is unreachable.
                </Text>
              ) : (
                <Text className={styles.rowValue}>
                  Traffic is allowed direct if the GSA edge is unreachable.
                </Text>
              )}
            </div>
          </div>

          {!isPrivate && (profile.policies ?? []).some((l) => (l.policy?.policyRules ?? []).length > 0) && (
            <div className={styles.serviceGroups}>
              {(profile.policies ?? [])
                .filter((l) => (l.policy?.policyRules ?? []).length > 0)
                .map((link) => {
                  const groupRules = link.policy?.policyRules ?? [];
                  const fwdCount = groupRules.filter((r) => r.action === 'forward').length;
                  const bypCount = groupRules.filter((r) => r.action === 'bypass').length;
                  const groupDisabled = link.state === 'disabled';
                  const { color: stateColor, Icon: StIcon, label: stateLabel } = stateBadge(groupDisabled ? 'disabled' : 'enabled');
                  const groupExpanded = expandedGroups.has(link.id);
                  const toggleGroup = () =>
                    setExpandedGroups((prev) => {
                      const next = new Set(prev);
                      if (groupExpanded) next.delete(link.id);
                      else next.add(link.id);
                      return next as ReadonlySet<string>;
                    });
                  const GroupChevron = groupExpanded ? ChevronDown16Regular : ChevronRight16Regular;
                  return (
                    <div key={link.id} className={styles.sgCard}>
                      <div
                        className={mergeClasses(styles.sgHeader, groupDisabled && styles.sgHeaderDisabled)}
                        onClick={toggleGroup}
                        role="button"
                        aria-expanded={groupExpanded}
                      >
                        <div className={styles.sgGutter}>
                          {link.priority != null && (
                            <>
                              <Text className={styles.sgGutterLabel}>Priority</Text>
                              <Text className={styles.sgGutterValue}>{link.priority}</Text>
                            </>
                          )}
                        </div>
                        <div className={styles.sgIdentity}>
                          <div className={styles.sgIdentityLead}>
                            <span className={styles.chevron}><GroupChevron /></span>
                            <Text className={styles.sgName}>{link.policy?.name ?? 'Policy'}</Text>
                          </div>
                          {groupRules.length > 0 && (
                            <Text className={styles.sgCounts}>
                              {[
                                fwdCount > 0 ? `${fwdCount} forward` : '',
                                bypCount > 0 ? `${bypCount} bypass` : '',
                              ].filter(Boolean).join(' · ')}
                            </Text>
                          )}
                        </div>
                        <div className={styles.sgAction}>
                          {type === 'm365' && (
                            <Badge appearance="tint" color={stateColor} icon={<StIcon />}>{stateLabel}</Badge>
                          )}
                        </div>
                      </div>
                      {groupExpanded && (
                        <div className={styles.ruleList}>
                          <div className={mergeClasses(styles.ruleGridRow, styles.ruleColTitles)}>
                            <span>Destination</span>
                            <span>Action</span>
                            <span>Type</span>
                          </div>
                          {groupRules.map((rule, rIdx) => {
                            const typeMeta = FWDRULE_TYPE_META[rule.ruleType ?? ''];
                            const TypeIcon = typeMeta?.Icon ?? Globe16Regular;
                            const typeColor: BadgeProps['color'] = typeMeta?.color ?? 'subtle';
                            return (
                              <div
                                key={rule.id}
                                className={mergeClasses(
                                  styles.ruleGridRow,
                                  styles.ruleRow,
                                  rIdx % 2 === 1 && styles.ruleRowAlt,
                                )}
                              >
                                <div className={styles.ruleCell}>
                                  <Text className={styles.ruleDest}>{rulePrimaryDest(rule)}</Text>
                                </div>
                                <div className={styles.ruleCell}>
                                  <Badge
                                    appearance="tint"
                                    color={rule.action === 'forward' ? 'success' : 'warning'}
                                    size="small"
                                  >
                                    {rule.action ?? '?'}
                                  </Badge>
                                </div>
                                <div className={styles.ruleCell}>
                                  {rule.ruleType && (
                                    <Badge
                                      appearance="tint"
                                      color={typeColor}
                                      icon={<TypeIcon />}
                                      size="small"
                                    >
                                      {rule.ruleType}
                                    </Badge>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}
            </div>
          )}

          {isPrivate && inProfileApps.length > 0 && (
            <div className={styles.row}>
              <Text className={styles.rowLabel}>Apps</Text>
              <div className={styles.appList}>
                {inProfileApps.map((app) => (
                  <Badge key={app.appId ?? app.id} appearance="outline" color="informative">
                    {app.displayName ?? app.appId ?? app.id}
                  </Badge>
                ))}
              </div>
            </div>
          )}

          {excludedApps.length > 0 && (
            <div className={styles.row}>
              <Text className={styles.rowLabel}>Not in profile</Text>
              <div className={styles.appList}>
                {excludedApps.map((app) => (
                  <Badge key={app.appId ?? app.id} appearance="outline" color="subtle">
                    {app.displayName ?? app.appId ?? app.id}
                  </Badge>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Public component
// ---------------------------------------------------------------------------

export interface ForwardingProfilesCardProps {
  profiles: ForwardingProfile[];
  types: TrafficForwardingType[];
  paApps?: PrivateAccessApp[];
  /** Show the traffic-type label in each card header.
   *  Defaults to true when multiple types are displayed, false when only one type
   *  (the enclosing section already provides context). Pass explicitly to override. */
  showTypeLabel?: boolean;
}

export function ForwardingProfilesCard({
  profiles,
  types,
  paApps = [],
  showTypeLabel,
}: ForwardingProfilesCardProps) {
  const styles = useStyles();

  const typeSet = new Set(types);
  const filtered = profiles
    .filter((p) => typeSet.has(p.trafficForwardingType ?? ''))
    .sort((a, b) => {
      const ti = types.indexOf(a.trafficForwardingType ?? '');
      const tj = types.indexOf(b.trafficForwardingType ?? '');
      if (ti !== tj) return ti - tj;
      return (a.priority ?? 999) - (b.priority ?? 999);
    });

  if (filtered.length === 0) return null;

  // Default: show type label only when the card spans multiple traffic types.
  const resolvedShowTypeLabel = showTypeLabel ?? (new Set(filtered.map((p) => p.trafficForwardingType)).size > 1);

  return (
    <div className={styles.list}>
      {filtered.map((profile) => (
        <ProfileCard key={profile.id} profile={profile} paApps={paApps} showTypeLabel={resolvedShowTypeLabel} />
      ))}
    </div>
  );
}
