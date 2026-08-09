import { useRef, useState, type ChangeEvent } from 'react';
import {
  Button,
  Divider,
  Field,
  Input,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useAuth } from '../../auth/authContext';
import { useDataSource } from '../../app/dataSourceContext';
import { DEFAULT_TENANT } from '../../auth/msalConfig';
import { FixtureParseError } from '../../adapters/file/fixture';
import { friendlyError } from '../friendlyError';

const useStyles = makeStyles({
  root: {
    display: 'grid',
    placeItems: 'center',
    minHeight: '100vh',
    padding: tokens.spacingVerticalL,
    boxSizing: 'border-box',
    // Soft neutral page like the Microsoft sign-in background.
    background:
      'radial-gradient(1200px 600px at 50% -10%, #eaf1fb 0%, #f3f3f3 55%, #ededed 100%)',
  },
  card: {
    width: '440px',
    maxWidth: '100%',
    boxSizing: 'border-box',
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalXXL}`,
    backgroundColor: tokens.colorNeutralBackground1,
    // Microsoft sign-in card: flat, square-ish, soft drop shadow.
    boxShadow: '0 2px 6px rgba(0, 0, 0, 0.18), 0 0 1px rgba(0, 0, 0, 0.12)',
    borderRadius: tokens.borderRadiusSmall,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
  },
  logoRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  logoText: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  heading: {
    fontSize: tokens.fontSizeBase600,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  muted: { color: tokens.colorNeutralForeground3 },
  // Square, Microsoft-blue primary button.
  primary: {
    backgroundColor: '#0067b8',
    borderRadius: 0,
    minWidth: '108px',
    alignSelf: 'center',
    ':hover': { backgroundColor: '#005da6' },
    ':hover:active': { backgroundColor: '#00528f' },
  },
  fileBtn: { borderRadius: 0 },
});

/** The Microsoft four-square logo (inline, no asset dependency). */
function MicrosoftLogo() {
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true" focusable="false">
      <rect x="1" y="1" width="8.5" height="8.5" fill="#F25022" />
      <rect x="10.5" y="1" width="8.5" height="8.5" fill="#7FBA00" />
      <rect x="1" y="10.5" width="8.5" height="8.5" fill="#00A4EF" />
      <rect x="10.5" y="10.5" width="8.5" height="8.5" fill="#FFB900" />
    </svg>
  );
}

export function SignInGate() {
  const styles = useStyles();
  const { signIn } = useAuth();
  const { enterFileMode } = useDataSource();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [tenant, setTenant] = useState(DEFAULT_TENANT);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function handleSignIn() {
    setBusy(true);
    setError(null);
    try {
      await signIn(tenant || 'organizations');
    } catch (e) {
      setError(friendlyError(e, 'Sign-in failed.'));
    } finally {
      setBusy(false);
    }
  }

  async function handleFileChosen(e: ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = ''; // allow re-selecting the same file later
    if (!file) return;
    setError(null);
    try {
      const text = await file.text();
      enterFileMode(text, file.name);
    } catch (err) {
      // The file name is deliberately not echoed back: exported fixture names
      // carry the tenant domain by construction (finding 15).
      if (err instanceof FixtureParseError) {
        setError(
          'The selected file doesn’t match the expected format. Choose a JSON ' +
            'file created by the exporter (testharness/Export-GsaFixture.ps1). ' +
            `See testharness/README.md for how to generate one.\n\nDetails: ${err.message}`,
        );
      } else {
        setError(friendlyError(err, 'That file could not be loaded.'));
      }
    }
  }

  return (
    <div className={styles.root}>
      <div className={styles.card}>
        <div className={styles.logoRow}>
          <MicrosoftLogo />
          <span className={styles.logoText}>Microsoft</span>
        </div>
        <div className={styles.heading}>Sign in</div>
        <Text className={styles.muted}>
          Sign in to a tenant with Global Secure Access Internet Access
          configured. You will be asked to consent to{' '}
          <code>NetworkAccess.Read.All</code>; the optional{' '}
          <code>Policy.Read.All</code> and <code>Directory.Read.All</code> add
          Conditional Access targeting detail with friendly names.
        </Text>
        <Field
          label="Tenant"
          hint="GUID, verified domain, or 'organizations' to pick at sign-in."
        >
          <Input
            value={tenant}
            onChange={(_, d) => setTenant(d.value.trim())}
            onFocus={() => {
              // Clear the pre-filled default the moment the field is focused so
              // a pasted domain doesn't get appended to it.
              if (tenant === DEFAULT_TENANT) setTenant('');
            }}
            placeholder="organizations"
          />
        </Field>
        <Button
          appearance="primary"
          className={styles.primary}
          disabled={busy}
          onClick={handleSignIn}
        >
          {busy ? 'Signing in…' : 'Sign in'}
        </Button>
        {error && (
          <Text
            size={200}
            style={{
              color: tokens.colorPaletteRedForeground1,
              whiteSpace: 'pre-line',
              backgroundColor: tokens.colorPaletteRedBackground1,
              border: `1px solid ${tokens.colorPaletteRedBorder1}`,
              borderRadius: tokens.borderRadiusMedium,
              padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
            }}
          >
            {error}
          </Text>
        )}
        <Divider>or</Divider>
        <input
          ref={fileInputRef}
          type="file"
          accept="application/json,.json"
          style={{ display: 'none' }}
          onChange={handleFileChosen}
        />
        <Button
          appearance="secondary"
          className={styles.fileBtn}
          onClick={() => fileInputRef.current?.click()}
        >
          Load policy file (no sign-in)
        </Button>
        <Text size={200} className={styles.muted}>
          Inspect an exported tenant configuration offline — no tenant or sign-in
          required.
        </Text>
      </div>
    </div>
  );
}
