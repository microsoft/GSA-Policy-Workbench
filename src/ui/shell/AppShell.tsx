import type { ReactNode } from 'react';
import {
  Badge,
  Button,
  Text,
  Tooltip,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ShieldTask20Regular,
  SignOut20Regular,
  Document20Regular,
  ArrowClockwise20Regular,
  ArrowDownload20Regular,
} from '@fluentui/react-icons';
import { useAuth } from '../../auth/authContext';
import { useDataSource } from '../../app/dataSourceContext';
import { SourceRail } from './SourceRail';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    boxSizing: 'border-box',
  },
  topbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    fontWeight: tokens.fontWeightSemibold,
  },
  spacer: { flex: 1 },
  muted: { color: tokens.colorNeutralForeground3 },
  // Row below topbar: SourceRail on the left, scrollable content on the right.
  body: {
    flex: 1,
    minHeight: 0,
    display: 'flex',
    flexDirection: 'row',
  },
  content: {
    flex: 1,
    minHeight: 0,
    overflowY: 'auto',
    padding: tokens.spacingHorizontalL,
    boxSizing: 'border-box',
  },
});

interface AppShellProps {
  children: ReactNode;
  /** Called when the user clicks Reload (live mode only). */
  onReload?: () => void;
  /** True while a reload fetch is in progress — spins the reload button. */
  isReloading?: boolean;
  /** Called when the user clicks Export (live mode only, data must be loaded). */
  onExport?: () => void;
  /** Branch note from the loader (e.g. load strategy) shown in the account tooltip. */
  connectionNote?: string;
}

export function AppShell({ children, onReload, isReloading, onExport, connectionNote }: AppShellProps) {
  const styles = useStyles();
  const { account, tenant, hasCaDetailScope, hasAllOptionalScopes, grantOptionalScopes, signOut } = useAuth();
  const { mode, fileName, exitFileMode } = useDataSource();

  const fileMode = mode === 'file';

  return (
    <div className={styles.root}>
      <div className={styles.topbar}>
        <span className={styles.brand}>
          <ShieldTask20Regular />
          GSA Policy Workbench
        </span>
        <div className={styles.spacer} />
        {fileMode ? (
          <>
            <Badge
              appearance="tint"
              color="informative"
              icon={<Document20Regular />}
            >
              File mode{fileName ? `: ${fileName}` : ''}
            </Badge>
            <Button
              appearance="subtle"
              icon={<SignOut20Regular />}
              onClick={exitFileMode}
            >
              Exit file mode
            </Button>
          </>
        ) : (
          <>
            <Tooltip
              content={
                <>
                  {hasCaDetailScope
                    ? 'Policy.Read.All granted — full where-used detail available.'
                    : 'Policy.Read.All not granted — where-used shows name + ID only.'}
                  {connectionNote && <><br />{connectionNote}</>}
                </>
              }
              relationship="description"
            >
              <Text as="span" size={200} className={styles.muted} tabIndex={0}>
                {account?.username ?? ''} · {tenant}
              </Text>
            </Tooltip>
            {!hasAllOptionalScopes && (
              <Tooltip
                content="Grant the optional read permissions (Conditional Access detail, directory names, Private Access segments, What-If user lookup). Requested only when you ask for them."
                relationship="label"
              >
                <Button
                  appearance="subtle"
                  onClick={() => void grantOptionalScopes()}
                >
                  Enable full detail
                </Button>
              </Tooltip>
            )}
            {onReload && (
              <Tooltip content="Reload data from Microsoft Graph" relationship="label">
                <Button
                  appearance="subtle"
                  icon={<ArrowClockwise20Regular />}
                  disabled={isReloading}
                  onClick={onReload}
                >
                  {isReloading ? 'Reloading…' : 'Reload'}
                </Button>
              </Tooltip>
            )}
            {onExport && (
              <Tooltip content="Export the loaded config as a fixture JSON file you can reload later. Contains real tenant data — do not commit." relationship="label">
                <Button
                  appearance="subtle"
                  icon={<ArrowDownload20Regular />}
                  onClick={onExport}
                >
                  Export
                </Button>
              </Tooltip>
            )}
            <Button
              appearance="subtle"
              icon={<SignOut20Regular />}
              onClick={() => void signOut()}
            >
              Sign out
            </Button>
          </>
        )}
      </div>
      <div className={styles.body}>
        <SourceRail />
        <div className={styles.content}>{children}</div>
      </div>
    </div>
  );
}
