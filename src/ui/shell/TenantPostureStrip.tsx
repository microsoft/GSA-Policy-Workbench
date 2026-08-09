/**
 * Tenant posture strip — Tier 1 (UI), read-only.
 *
 * A compact "tenant posture" band shown atop the Internet Access view, beneath
 * the forwarding-profile strip. Surfaces two V2 read-only intelligence signals
 * (spec §10):
 *
 *   - Tenant-limit counters: current usage as "N / MAX" chips against the
 *     published Internet Access limits, neutral until 80% (amber) / 100% (red).
 *   - Orphaned-policy hint: a collapsed chip ("N policies not linked to any
 *     profile") that expands to a list grouped by policy type.
 *
 * Uses the shared posture vocabulary (amber = review, neutral = informational).
 * These are inspection hints, never a pass/fail verdict. Display only — the
 * computation is pure (tenantPosture.ts) and no Graph calls happen here.
 */

import { useMemo, useState } from 'react';
import {
  Badge,
  Text,
  Tooltip,
  makeStyles,
  mergeClasses,
  tokens,
  type BadgeProps,
} from '@fluentui/react-components';
import {
  ChevronDown16Regular,
  ChevronRight16Regular,
  PlugDisconnected16Regular,
  Info16Regular,
} from '@fluentui/react-icons';
import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type { TenantPolicy } from '../../adapters/graph/loader';
import { shortPolicyType } from '../policyTypeMeta';
import {
  computeTenantPosture,
  type CounterLevel,
  type LimitCounter,
} from './tenantPosture';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    flexShrink: 0,
    padding: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  label: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    textTransform: 'uppercase',
    letterSpacing: '0.03em',
  },
  grow: { flex: 1 },

  // counter chip
  counter: {
    display: 'inline-flex',
    alignItems: 'baseline',
    gap: tokens.spacingHorizontalXXS,
    paddingTop: tokens.spacingVerticalXXS,
    paddingBottom: tokens.spacingVerticalXXS,
    paddingLeft: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  counterWarn: {
    border: `1px solid ${tokens.colorPaletteDarkOrangeBorderActive}`,
    backgroundColor: tokens.colorPaletteDarkOrangeBackground1,
  },
  counterFull: {
    border: `1px solid ${tokens.colorPaletteRedBorderActive}`,
    backgroundColor: tokens.colorPaletteRedBackground1,
  },
  counterValue: { fontWeight: tokens.fontWeightBold },
  counterMax: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  counterLabel: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    marginLeft: tokens.spacingHorizontalXXS,
  },

  // orphan hint
  orphanHead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    cursor: 'pointer',
    userSelect: 'none',
  },
  chevron: { color: tokens.colorNeutralForeground2, display: 'inline-flex' },
  orphanList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    paddingLeft: tokens.spacingHorizontalL,
  },
  orphanGroup: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  orphanName: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
  },
  allLinked: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
});

const LEVEL_TIP: Record<CounterLevel, string> = {
  ok: '',
  warn: 'Approaching the tenant limit (≥ 80%).',
  full: 'At the tenant limit — new items may be rejected.',
};

function CounterChip({ counter }: { counter: LimitCounter }) {
  const styles = useStyles();
  const cls = mergeClasses(
    styles.counter,
    counter.level === 'warn' && styles.counterWarn,
    counter.level === 'full' && styles.counterFull,
  );
  const pct = Math.round((counter.count / counter.max) * 100);
  const tip = [
    `${counter.label}: ${counter.count} of ${counter.max} (${pct}%).`,
    counter.linkedOnly
      ? 'Counted across policies linked to a profile.'
      : '',
    LEVEL_TIP[counter.level],
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <Tooltip content={tip} relationship="label">
      <span className={cls}>
        <span className={styles.counterValue}>{counter.count.toLocaleString()}</span>
        <span className={styles.counterMax}>/ {counter.max.toLocaleString()}</span>
        <span className={styles.counterLabel}>{counter.label}</span>
      </span>
    </Tooltip>
  );
}

interface TenantPostureStripProps {
  profiles: SecurityProfile[];
  tenantPolicies: TenantPolicy[];
}

export function TenantPostureStrip({ profiles, tenantPolicies }: TenantPostureStripProps) {
  const styles = useStyles();
  const [open, setOpen] = useState(false);
  const posture = useMemo(
    () => computeTenantPosture(profiles, tenantPolicies),
    [profiles, tenantPolicies],
  );

  const hasInventory = tenantPolicies.length > 0;

  return (
    <div className={styles.root}>
      <div className={styles.row}>
        <Text className={styles.label}>Tenant posture</Text>
        {posture.counters.map((c) => (
          <CounterChip key={c.key} counter={c} />
        ))}
      </div>

      {hasInventory && (
        <OrphanHint
          posture={posture}
          open={open}
          onToggle={() => setOpen((v) => !v)}
        />
      )}
    </div>
  );
}

function OrphanHint({
  posture,
  open,
  onToggle,
}: {
  posture: ReturnType<typeof computeTenantPosture>;
  open: boolean;
  onToggle: () => void;
}) {
  const styles = useStyles();

  if (posture.orphanCount === 0) {
    return (
      <span className={styles.allLinked}>
        <Info16Regular />
        Every tenant policy is linked to a profile.
      </span>
    );
  }

  const hintColor: BadgeProps['color'] = 'warning';

  return (
    <div>
      <div
        className={styles.orphanHead}
        role="button"
        tabIndex={0}
        aria-expanded={open}
        onClick={onToggle}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            onToggle();
          }
        }}
      >
        <span className={styles.chevron}>
          {open ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
        </span>
        <Tooltip
          content="These policies exist in the tenant but are not linked to any Security Profile, so they are not enforced. Inspection hint only."
          relationship="label"
        >
          <Badge appearance="tint" color={hintColor} icon={<PlugDisconnected16Regular />}>
            {posture.orphanCount}{' '}
            {posture.orphanCount === 1 ? 'policy' : 'policies'} not linked to any profile
          </Badge>
        </Tooltip>
      </div>

      {open && (
        <div className={styles.orphanList}>
          {posture.orphaned.map((group) => (
            <div key={group.typeLabel} className={styles.orphanGroup}>
              <Badge appearance="tint" color={group.color}>
                {shortPolicyType(group.typeLabel)}
              </Badge>
              {group.policies.map((p) => (
                <span key={p.id} className={styles.orphanName}>
                  {p.name}
                </span>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
