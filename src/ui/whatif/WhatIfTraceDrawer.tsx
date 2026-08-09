/**
 * What-If trace drawer — on-demand decision-tree view.
 *
 * Renders a structured step-by-step explanation of how the What-If resolver
 * reached its outcome. Opened by the "Explain" button in WhatIfPanel.
 * Receives a pre-computed TraceRecord[] — no new evaluation logic here.
 */

import {
  Badge,
  Button,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  CheckmarkCircle16Regular,
  DismissCircle16Regular,
  Dismiss24Regular,
  FlashCheckmark16Regular,
  Info16Regular,
  Warning16Regular,
  ArrowUpRight16Regular,
  RecordStop16Regular,
  Trophy16Regular,
} from '@fluentui/react-icons';
import type { TraceDecision, TraceRecord } from '../../app/tracer';

const STAGE_LABEL: Record<string, string> = {
  context:      'Resolved context',
  acquisition: 'Stage 1 — Traffic acquisition',
  profile:     'Stage 2 — Profile selection',
  rule:        'Stage 3 — Rule evaluation',
  'pa-segment': 'Private Access segments',
};

const STAGE_ORDER = ['context', 'acquisition', 'profile', 'rule', 'pa-segment'];

const useStyles = makeStyles({
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalL}`,
    overflowY: 'auto',
  },
  empty: {
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
    paddingBottom: tokens.spacingVerticalXS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  row: {
    display: 'grid',
    gridTemplateColumns: '20px 1fr',
    gap: tokens.spacingHorizontalS,
    alignItems: 'start',
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalXS}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  rowWinner: {
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke2}`,
  },
  rowEscape: {
    backgroundColor: tokens.colorPaletteDarkOrangeBackground1,
    border: `1px solid ${tokens.colorPaletteDarkOrangeBorderActive}`,
  },
  rowSkip: { opacity: 0.6 },
  icon: { paddingTop: '2px', display: 'flex' },
  content: { display: 'flex', flexDirection: 'column', gap: '2px' },
  label: { fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 },
  reason: { color: tokens.colorNeutralForeground2, fontSize: tokens.fontSizeBase200 },
  badges: {
    display: 'flex',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
    marginTop: '2px',
  },
});

function DecisionIcon({ decision }: { decision: TraceDecision }) {
  switch (decision) {
    case 'winner':    return <Trophy16Regular style={{ color: tokens.colorBrandForeground1 }} />;
    case 'pass':      return <CheckmarkCircle16Regular style={{ color: tokens.colorPaletteGreenForeground1 }} />;
    case 'match':     return <FlashCheckmark16Regular style={{ color: tokens.colorBrandForeground1 }} />;
    case 'skip':      return <DismissCircle16Regular style={{ color: tokens.colorNeutralForeground4 }} />;
    case 'no-match':  return <RecordStop16Regular style={{ color: tokens.colorNeutralForeground4 }} />;
    case 'escape':    return <Warning16Regular style={{ color: tokens.colorPaletteDarkOrangeForeground1 }} />;
    case 'preempted': return <ArrowUpRight16Regular style={{ color: tokens.colorPaletteBlueForeground2 }} />;
    default:          return <Info16Regular style={{ color: tokens.colorNeutralForeground3 }} />;
  }
}

function DetailBadges({ detail }: { detail?: Record<string, unknown> }) {
  const styles = useStyles();
  if (!detail) return null;
  const entries = Object.entries(detail).filter(
    ([, v]) => v !== undefined && v !== null && v !== '' && v !== '(none)' && !Array.isArray(v),
  );
  if (entries.length === 0) return null;
  return (
    <div className={styles.badges}>
      {entries.map(([k, v]) => (
        <Badge key={k} appearance="tint" color="subtle" size="small">
          {k}: {String(v)}
        </Badge>
      ))}
    </div>
  );
}

function TraceRow({ record }: { record: TraceRecord }) {
  const styles = useStyles();
  const isWinner = record.decision === 'winner';
  const isEscape = record.decision === 'escape';
  const isSkip = record.decision === 'skip' || record.decision === 'no-match';
  return (
    <div
      className={`${styles.row} ${isWinner ? styles.rowWinner : ''} ${isEscape ? styles.rowEscape : ''} ${isSkip ? styles.rowSkip : ''}`}
    >
      <span className={styles.icon}>
        <DecisionIcon decision={record.decision} />
      </span>
      <div className={styles.content}>
        <Text className={styles.label}>{record.label}</Text>
        <Text className={styles.reason}>{record.reason}</Text>
        <DetailBadges detail={record.detail} />
      </div>
    </div>
  );
}

interface WhatIfTraceDrawerProps {
  open: boolean;
  onClose: () => void;
  records: TraceRecord[];
}

export function WhatIfTraceDrawer({ open, onClose, records }: WhatIfTraceDrawerProps) {
  const styles = useStyles();

  const stages = STAGE_ORDER.filter((s) => records.some((r) => r.stage === s));

  return (
    <OverlayDrawer
      position="end"
      size="medium"
      open={open}
      onOpenChange={(_, { open: o }) => { if (!o) onClose(); }}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close trace"
              icon={<Dismiss24Regular />}
              onClick={onClose}
            />
          }
        >
          Decision trace
        </DrawerHeaderTitle>
      </DrawerHeader>

      <DrawerBody>
        <div className={styles.body}>
          {records.length === 0 ? (
            <Text className={styles.empty}>
              Enter a user or destination in the What-If panel to see the decision trace.
            </Text>
          ) : (
            stages.map((stage) => (
              <div key={stage} className={styles.section}>
                <Text className={styles.sectionTitle}>
                  {STAGE_LABEL[stage] ?? stage}
                </Text>
                {records
                  .filter((r) => r.stage === stage)
                  .map((r, i) => (
                    <TraceRow key={`${stage}-${i}`} record={r} />
                  ))}
              </div>
            ))
          )}
        </div>
      </DrawerBody>
    </OverlayDrawer>
  );
}
