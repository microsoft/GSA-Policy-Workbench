import type { ReactNode } from 'react';
import {
  Button,
  Card,
  Text,
  Title2,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useAuth } from '../../auth/authContext';

const useStyles = makeStyles({
  root: {
    display: 'grid',
    placeItems: 'center',
    height: '100vh',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  card: {
    width: '480px',
    padding: tokens.spacingVerticalXXL,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  muted: { color: tokens.colorNeutralForeground3 },
});

/**
 * GSA Internet Access Graph APIs are available only in the Global service
 * deployment, so government and sovereign tenants are blocked here.
 *
 * `tenant_region_scope` is an *optional* claim, absent unless the app
 * registration emits it — an allow-list on `WW` would lock out every ordinary
 * commercial tenant. The hard sovereign boundary is already the authority: this
 * app only talks to login.microsoftonline.com, which US Gov
 * (login.microsoftonline.us) and 21Vianet (login.partner.microsoftonline.cn)
 * cannot use. So this guard blocks on a *positive* government signal, which is
 * what remains reachable — chiefly GCC tenants hosted in the commercial cloud.
 */
const GOV_REGION_SCOPES = new Set(['usgov', 'usg', 'usnat', 'ussec', 'dod', 'cn']);
const GOV_SUB_SCOPES = new Set(['gcc', 'gcch', 'dod', 'dodcon']);

function claimText(claims: Record<string, unknown>, key: string): string {
  const v = claims[key];
  return typeof v === 'string' ? v.trim().toLowerCase() : '';
}

export function GovCloudGuard({ children }: { children: ReactNode }) {
  const styles = useStyles();
  const { account, signOut } = useAuth();

  const claims = (account?.idTokenClaims ?? {}) as Record<string, unknown>;
  const regionScope = claimText(claims, 'tenant_region_scope');
  const subScope = claimText(claims, 'tenant_region_sub_scope');

  if (!account || GOV_REGION_SCOPES.has(regionScope) || GOV_SUB_SCOPES.has(subScope)) {
    return (
      <div className={styles.root}>
        <Card className={styles.card}>
          <Title2>Global service only</Title2>
          <Text className={styles.muted}>
            This tool is only available for commercial (worldwide) tenants.
            Microsoft Entra Government and 21Vianet tenants are not supported.
            GSA Internet Access policy APIs exist only in the Global service
            deployment, so the Workbench cannot operate against this tenant.
          </Text>
          <Button appearance="primary" onClick={() => void signOut()}>
            Sign out
          </Button>
        </Card>
      </div>
    );
  }

  return <>{children}</>;
}
