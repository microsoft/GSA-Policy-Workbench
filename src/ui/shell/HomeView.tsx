/**
 * Home landing view — Tier 1 (UI).
 *
 * The entry screen once a data source is active. Visual language is aligned
 * with the GSA tooling family (the Private Access Sizing Planner): a gradient
 * hero band with capability pills, followed by "choose your domain" mode cards.
 *
 *   • Internet Access (IA) — fully supported; opens the policy table.
 *   • Private Access (PA)  — read-only inspection of PA / App Proxy apps and
 *     their Conditional Access correlation. Enabled when the active data source
 *     carries PA data (file mode today; spec §6), otherwise a disabled card
 *     explains that no PA data is loaded.
 *
 * Display only — no Graph calls.
 */

import {
  Badge,
  Button,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import {
  Globe24Regular,
  LockClosed24Regular,
  CheckmarkCircle16Filled,
  ArrowRight16Regular,
  ShieldCheckmark16Regular,
  DocumentSearch16Regular,
  Eye16Regular,
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXL,
    width: '100%',
    maxWidth: '1080px',
    marginLeft: 'auto',
    marginRight: 'auto',
  },

  // ---- Hero band (planner-inspired gradient header) --------------------
  heroBand: {
    width: '100%',
    maxWidth: '1080px',
    marginLeft: 'auto',
    marginRight: 'auto',
  },
  hero: {
    position: 'relative',
    overflow: 'hidden',
    borderRadius: tokens.borderRadiusXLarge,
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalXXL}`,
    color: tokens.colorNeutralForegroundOnBrand,
    background: 'linear-gradient(135deg, #0b2a6b 0%, #1f4fb0 45%, #5a3bb8 100%)',
    boxShadow: tokens.shadow8,
    textAlign: 'center',
  },
  heroEyebrow: {
    textTransform: 'uppercase',
    letterSpacing: '0.08em',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    opacity: 0.85,
  },
  heroTitle: {
    fontSize: tokens.fontSizeHero800,
    lineHeight: tokens.lineHeightHero800,
    fontWeight: tokens.fontWeightBold,
    margin: `${tokens.spacingVerticalS} 0`,
  },
  heroSubtitle: {
    fontSize: tokens.fontSizeBase400,
    maxWidth: '64ch',
    marginLeft: 'auto',
    marginRight: 'auto',
    opacity: 0.92,
  },
  pills: {
    display: 'flex',
    flexWrap: 'wrap',
    justifyContent: 'center',
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalL,
  },
  pill: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    padding: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: 'rgba(255, 255, 255, 0.16)',
    border: '1px solid rgba(255, 255, 255, 0.28)',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },

  // ---- Mode cards -----------------------------------------------------
  sectionLabel: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
  },
  cards: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))',
    gap: tokens.spacingHorizontalL,
  },
  card: {
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalXL,
    borderRadius: tokens.borderRadiusXLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
    overflow: 'hidden',
  },
  cardAccent: {
    position: 'absolute',
    insetBlockStart: 0,
    insetInlineStart: 0,
    insetInlineEnd: 0,
    height: '4px',
    background: 'linear-gradient(90deg, #1f4fb0, #5a3bb8)',
  },
  cardDisabled: {
    backgroundColor: tokens.colorNeutralBackground2,
    boxShadow: 'none',
  },
  cardHead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
  },
  cardIcon: {
    display: 'grid',
    placeItems: 'center',
    width: '48px',
    height: '48px',
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    flexShrink: 0,
  },
  cardIconMuted: { backgroundColor: tokens.colorNeutralBackground4, color: tokens.colorNeutralForeground3 },
  cardTitle: { fontSize: tokens.fontSizeBase500, fontWeight: tokens.fontWeightBold },
  cardKicker: { color: tokens.colorNeutralForeground3, fontWeight: tokens.fontWeightSemibold },
  cardDesc: { color: tokens.colorNeutralForeground2 },
  bullets: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    margin: 0,
    padding: 0,
    listStyle: 'none',
  },
  bullet: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: tokens.colorNeutralForeground2,
  },
  bulletIcon: { color: tokens.colorBrandForeground1, flexShrink: 0 },
  bulletIconMuted: { color: tokens.colorNeutralForeground4, flexShrink: 0 },
  cardFoot: {
    marginTop: 'auto',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
  },
});

interface HomeViewProps {
  /** Number of Internet Access security profiles (for the IA card stat). */
  iaProfileCount: number;
  /** Number of Private Access + App Proxy apps available in the data source. */
  paAppCount: number;
  onOpenIa: () => void;
  onOpenPa: () => void;
}

const HERO_PILLS = [
  { icon: <Eye16Regular />, label: 'Read-only inspection' },
  { icon: <ShieldCheckmark16Regular />, label: 'Conditional Access aware' },
  { icon: <DocumentSearch16Regular />, label: 'Every Graph call audited' },
];

/**
 * The "Policy Workbench" hero band. Rendered at the very top of the home view,
 * above the What-If band, so the page leads with the product identity.
 */
export function HomeHero() {
  const styles = useStyles();
  return (
    <div className={styles.heroBand}>
      <div className={styles.hero}>
        <div className={styles.heroEyebrow}>Microsoft Global Secure Access</div>
        <div className={styles.heroTitle}>Policy Workbench</div>
        <Text className={styles.heroSubtitle}>
          Inspect your Global Secure Access posture in one place — Internet Access
          Security Profiles and filtering rules, Private Access applications, and
          the Conditional Access policies that scope them to your users.
        </Text>
        <div className={styles.pills}>
          {HERO_PILLS.map((p) => (
            <span key={p.label} className={styles.pill}>
              {p.icon}
              {p.label}
            </span>
          ))}
        </div>
      </div>
    </div>
  );
}

export function HomeView({ iaProfileCount, paAppCount, onOpenIa, onOpenPa }: HomeViewProps) {
  const styles = useStyles();
  const paEnabled = paAppCount > 0;

  return (
    <div className={styles.root}>
      <Text className={styles.sectionLabel}>Choose a policy domain</Text>

      <div className={styles.cards}>
        {/* Internet Access — supported */}
        <div className={styles.card}>
          <span className={styles.cardAccent} />
          <div className={styles.cardHead}>
            <span className={styles.cardIcon}>
              <Globe24Regular />
            </span>
            <div>
              <Text className={styles.cardKicker}>Internet Access</Text>
              <div className={styles.cardTitle}>IA Policies</div>
            </div>
          </div>
          <Text className={styles.cardDesc}>
            Inspect Security Profiles, web content filtering policies, and their rules,
            with the linked Conditional Access targeting.
          </Text>
          <ul className={styles.bullets}>
            {[
              'Security Profiles by priority',
              'Filtering policies & rules',
              'Conditional Access targeting',
              'Search, filter, and drill down',
            ].map((b) => (
              <li key={b} className={styles.bullet}>
                <CheckmarkCircle16Filled className={styles.bulletIcon} />
                {b}
              </li>
            ))}
          </ul>
          <div className={styles.cardFoot}>
            <Badge appearance="tint" color="brand">
              {iaProfileCount} profile{iaProfileCount === 1 ? '' : 's'}
            </Badge>
            <Button
              appearance="primary"
              iconPosition="after"
              icon={<ArrowRight16Regular />}
              onClick={onOpenIa}
            >
              Open IA Policies
            </Button>
          </div>
        </div>

        {/* Private Access — enabled when data is present */}
        <div className={mergeClasses(styles.card, !paEnabled && styles.cardDisabled)}>
          {paEnabled && <span className={styles.cardAccent} />}
          <div className={styles.cardHead}>
            <span className={mergeClasses(styles.cardIcon, !paEnabled && styles.cardIconMuted)}>
              <LockClosed24Regular />
            </span>
            <div>
              <Text className={styles.cardKicker}>Private Access</Text>
              <div className={styles.cardTitle}>PA Policies</div>
            </div>
          </div>
          <Text className={styles.cardDesc}>
            Private Access apps, Application Proxy apps, and their Conditional Access
            correlation — including pre-authentication posture.
          </Text>
          <ul className={styles.bullets}>
            {[
              'Private Access & Quick Access apps',
              'Application Proxy pre-auth posture',
              'Per-app Conditional Access coverage',
            ].map((b) => (
              <li key={b} className={styles.bullet}>
                <CheckmarkCircle16Filled
                  className={paEnabled ? styles.bulletIcon : styles.bulletIconMuted}
                />
                {b}
              </li>
            ))}
          </ul>
          <div className={styles.cardFoot}>
            {paEnabled ? (
              <>
                <Badge appearance="tint" color="brand">
                  {paAppCount} app{paAppCount === 1 ? '' : 's'}
                </Badge>
                <Button
                  appearance="primary"
                  iconPosition="after"
                  icon={<ArrowRight16Regular />}
                  onClick={onOpenPa}
                >
                  Open PA Policies
                </Button>
              </>
            ) : (
              <>
                <Badge appearance="outline" color="informative">
                  No PA data loaded
                </Badge>
                <Button appearance="secondary" disabled>
                  Unavailable
                </Button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
