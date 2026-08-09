/**
 * Forwarding-profile status strip — Tier 1 (UI), read-only.
 *
 * Surfaces the GSA traffic-forwarding profile state (spec §6.6) as a compact
 * row of chips — the admin's first-order "is GSA on for this traffic class?"
 * signal. Shown atop each domain view: Internet (+ Microsoft 365) on the
 * Internet Access view, Private Access on the Private Access view.
 *
 * Uses the shared posture vocabulary (item 6): green = healthy / on,
 * amber = something to review (disabled), neutral = informational (not
 * configured). Display only — no Graph calls.
 */

import { Badge, Text, Tooltip, makeStyles, tokens, type BadgeProps } from '@fluentui/react-components';
import {
  CheckmarkCircle16Filled,
  Warning16Filled,
  Info16Regular,
} from '@fluentui/react-icons';
import type {
  ForwardingProfile,
  TrafficForwardingType,
} from '../../model/definitions/ForwardingProfile.definition';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
    flexShrink: 0,
  },
  label: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    textTransform: 'uppercase',
    letterSpacing: '0.03em',
  },
});

const TYPE_LABEL: Record<string, string> = {
  internet: 'Internet',
  m365: 'Microsoft 365',
  private: 'Private Access',
};

interface ChipMeta {
  color: BadgeProps['color'];
  icon: typeof CheckmarkCircle16Filled;
  suffix: string;
}

function chipMeta(state: string | undefined): ChipMeta {
  if (state === 'enabled') {
    return { color: 'success', icon: CheckmarkCircle16Filled, suffix: 'on' };
  }
  if (state === 'disabled') {
    return { color: 'warning', icon: Warning16Filled, suffix: 'disabled' };
  }
  return { color: 'subtle', icon: Info16Regular, suffix: 'not configured' };
}

/**
 * Summarise a profile's acquisition rules for the
 * chip tooltip — how many rules forward vs. explicitly bypass, and the
 * profile's evaluation priority. Absent when rule detail wasn't captured
 * (e.g. an older fixture) — the tooltip then just shows the priority, if any.
 */
function acquisitionSummary(profile: ForwardingProfile | undefined): string {
  if (!profile) return 'Not configured in this tenant.';
  const parts: string[] = [];
  if (profile.priority !== undefined) parts.push(`Priority ${profile.priority}`);
  const rules = (profile.policies ?? []).flatMap((link) => link.policy?.policyRules ?? []);
  if (rules.length > 0) {
    const bypassCount = rules.filter((r) => r.action === 'bypass').length;
    parts.push(
      `${rules.length} acquisition rule${rules.length === 1 ? '' : 's'}` +
        (bypassCount > 0 ? ` (${bypassCount} bypass)` : ''),
    );
  }
  return parts.length > 0 ? parts.join(' · ') : 'No acquisition-rule detail captured.';
}

interface ForwardingProfileStripProps {
  profiles: ForwardingProfile[];
  /** Which traffic classes to show, in order. */
  types: TrafficForwardingType[];
}

export function ForwardingProfileStrip({ profiles, types }: ForwardingProfileStripProps) {
  const styles = useStyles();
  const byType = new Map(profiles.map((p) => [p.trafficForwardingType, p]));

  // Nothing to show (e.g. live fetch returned nothing) — render nothing.
  if (profiles.length === 0) return null;

  return (
    <div className={styles.root}>
      <Text className={styles.label}>GSA forwarding</Text>
      {types.map((type) => {
        const profile = byType.get(type);
        const meta = chipMeta(profile?.state);
        const Icon = meta.icon;
        const label = TYPE_LABEL[type] ?? type;
        return (
          <Tooltip key={type} content={acquisitionSummary(profile)} relationship="description">
            <Badge appearance="tint" color={meta.color} icon={<Icon />}>
              {label} {meta.suffix}
            </Badge>
          </Tooltip>
        );
      })}
    </div>
  );
}
