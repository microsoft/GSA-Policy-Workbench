/**
 * Workbench — Tier 1 (UI), the main view once a data source is active.
 *
 * Renders Internet Access and Private Access on a single scrollable page,
 * each in a collapsible section card. The What-If band stays at the top
 * and applies across both domains.
 *
 * Covers loading / error / empty / data states.
 */

import { useMemo, useState } from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Spinner,
  Tab,
  TabList,
  Text,
  ToggleButton,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import {
  Globe20Regular,
  LockClosed20Regular,
  LockShield20Regular,
  Beaker20Regular,
  Tag20Regular,
  ChevronDown16Regular,
  ChevronRight16Regular,
} from '@fluentui/react-icons';
import { useProfileTree } from '../../query/hooks/useSecurityProfiles';
import { useDataSource } from '../../app/dataSourceContext';
import { useAuth } from '../../auth/authContext';
import { downloadFixture, fixtureFilename } from '../../adapters/file/exportFixture';
import { buildProfileGroups } from '../table/policyRows';
import { PolicyTable } from '../table/PolicyTable';
import { PrivateAccessView } from '../private/PrivateAccessView';
import { WhatIfPanel, WebCategoryLookup } from '../whatif/WhatIfPanel';
import { ForwardingProfilesCard } from './ForwardingProfilesCard';
import { TenantPostureStrip } from './TenantPostureStrip';
import { AppShell } from './AppShell';

const useStyles = makeStyles({
  center: {
    display: 'grid',
    placeItems: 'center',
    height: '100%',
    gap: tokens.spacingVerticalM,
    textAlign: 'center',
  },
  error: { color: tokens.colorPaletteRedForeground1 },

  // Outer column — natural height; outer AppShell content area scrolls.
  shell: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
  },

  // What-If band — centered, never shrinks.
  whatIfBand: {
    width: '100%',
    maxWidth: '1080px',
    marginLeft: 'auto',
    marginRight: 'auto',
    paddingTop: tokens.spacingVerticalL,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    flexShrink: 0,
  },
  whatIfHead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    cursor: 'pointer',
  },
  whatIfTitle: { fontSize: tokens.fontSizeBase500, fontWeight: tokens.fontWeightBold },
  whatIfSub: { color: tokens.colorNeutralForeground3 },
  whatIfChevron: { color: tokens.colorNeutralForeground2, display: 'inline-flex' },
  whatIfGrow: { flex: 1 },
  panelCard: {
    padding: tokens.spacingVerticalL,
    borderRadius: tokens.borderRadiusXLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
  },

  // Section cards (IA and PA).
  sectionCard: {
    // Base: flex column so head + body stack; no flex-grow yet (added conditionally).
    display: 'flex',
    flexDirection: 'column',
    borderRadius: tokens.borderRadiusXLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
    overflow: 'hidden',
  },
  // Applied when section is expanded — card grows to natural content height.
  sectionCardOpen: {},
  sectionHead: {
    flexShrink: 0,
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    cursor: 'pointer',
    userSelect: 'none',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  sectionHeadCollapsed: {
    borderBottom: 'none',
  },
  sectionHeadDisabled: {
    cursor: 'default',
    ':hover': { backgroundColor: 'unset' },
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  sectionGrow: { flex: 1 },
  sectionChevron: { color: tokens.colorNeutralForeground3, display: 'inline-flex' },

  // IA section body — grows to full content height (page scrolls).
  iaSectionBody: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalL,
  },

  // PA section body — grows to full content height (page scrolls).
  paSectionBody: {
    display: 'flex',
    flexDirection: 'column',
    padding: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalL,
  },

  // Section navigation bar — IA / PA toggle buttons.
  navBar: {
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },

  // Empty-state shown in PA section when no PA data is loaded.
  paEmptyState: {
    display: 'grid',
    placeItems: 'center',
    padding: `${tokens.spacingVerticalXXL} 0`,
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
    gap: tokens.spacingVerticalS,
  },
});

export function Workbench() {
  const styles = useStyles();
  const { activeSourceId } = useDataSource();
  const { account } = useAuth();
  const { data, isLoading, isError, error, refetch, isFetching } = useProfileTree(activeSourceId);
  const [diagOpen, setDiagOpen] = useState(true);
  const [diagTab, setDiagTab] = useState<'whatif' | 'category'>('whatif');
  const [m365Open, setM365Open] = useState(true);
  const [iaOpen, setIaOpen] = useState(true);
  const [paOpen, setPaOpen] = useState(true);
  const [exportOpen, setExportOpen] = useState(false);

  const groups = useMemo(
    () => (data ? buildProfileGroups(data.profiles, data.caDetails, data.directory) : []),
    [data],
  );

  const paAppCount =
    (data?.privateAccess.apps.length ?? 0) +
    (data?.privateAccess.appProxyApps.length ?? 0);
  const hasPaData = paAppCount > 0;
  const m365Profiles = (data?.forwardingProfiles ?? []).filter((p) => p.trafficForwardingType === 'm365');
  const hasMsData = m365Profiles.length > 0;
  const m365Count = m365Profiles.length;

  if (isLoading) {
    return (
      <AppShell>
        <div className={styles.center}>
          <Spinner label="Loading policy tree…" />
        </div>
      </AppShell>
    );
  }

  if (isError) {
    return (
      <AppShell>
        <div className={styles.center}>
          <Text className={styles.error}>
            Failed to load policies: {error?.message ?? 'Unknown error.'}
          </Text>
          <Button appearance="primary" onClick={() => void refetch()}>
            Retry
          </Button>
        </div>
      </AppShell>
    );
  }

  if (groups.length === 0) {
    return (
      <AppShell>
        <div className={styles.center}>
          <Text>No Security Profiles found in this data source.</Text>
        </div>
      </AppShell>
    );
  }

  return (
    <AppShell
      onReload={() => void refetch()}
      isReloading={isFetching}
      onExport={() => setExportOpen(true)}
      connectionNote={data?.branchNote}
    >
      <Dialog
        open={exportOpen}
        onOpenChange={(_, d) => setExportOpen(d.open)}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Export tenant configuration?</DialogTitle>
            <DialogContent>
              <Text as="p">
                This writes a JSON file containing the configuration currently
                loaded from your tenant: security profiles and rules,
                Conditional Access policy detail, resolved user and group
                display names, Private Access apps and their segment host names,
                forwarding profiles, and the tenant policy inventory.
              </Text>
              <Text as="p">
                It is real tenant data. Store it accordingly and never commit it
                to version control.
              </Text>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setExportOpen(false)}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => {
                  setExportOpen(false);
                  if (data) {
                    downloadFixture(
                      data,
                      fixtureFilename(account?.username ?? 'export'),
                    );
                  }
                }}
              >
                Export tenant data
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      <div className={styles.shell}>

        {/* ── Diagnostic Tools band ────────────────────────────────── */}
        <div className={styles.whatIfBand}>
          <div
            className={styles.whatIfHead}
            role="button"
            tabIndex={0}
            onClick={() => setDiagOpen((v) => !v)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                setDiagOpen((v) => !v);
              }
            }}
          >
            <Beaker20Regular />
            <Text className={styles.whatIfTitle}>Diagnostic Tools</Text>
            <span className={styles.whatIfGrow} />
            <span className={styles.whatIfChevron}>
              {diagOpen ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
            </span>
          </div>
          {diagOpen && (
            <div className={styles.panelCard}>
              <TabList
                selectedValue={diagTab}
                onTabSelect={(_, d) => setDiagTab(d.value as typeof diagTab)}
                appearance="subtle"
              >
                <Tab value="whatif" icon={<Beaker20Regular />}>Effective Policy</Tab>
                <Tab value="category" icon={<Tag20Regular />}>Category Lookup</Tab>
              </TabList>
              {diagTab === 'whatif' && (
                <WhatIfPanel
                  profiles={data?.profiles ?? []}
                  caDetails={data?.caDetails ?? []}
                  directory={data?.directory ?? []}
                  privateAccess={data?.privateAccess}
                  forwardingProfiles={data?.forwardingProfiles ?? []}
                  showWebCategory={false}
                />
              )}
              {diagTab === 'category' && <WebCategoryLookup />}
            </div>
          )}
        </div>

        {/* ── Section navigation bar ───────────────────────────────── */}
        <div className={styles.navBar}>
          <ToggleButton
            checked={iaOpen && paOpen && (!hasMsData || m365Open)}
            onClick={() => { setM365Open(true); setIaOpen(true); setPaOpen(true); }}
          >
            All
          </ToggleButton>
          <ToggleButton
            icon={<LockShield20Regular />}
            checked={m365Open && !paOpen && !iaOpen}
            disabled={!hasMsData}
            onClick={() => { setM365Open(true); setIaOpen(false); setPaOpen(false); }}
          >
            Microsoft 365{hasMsData ? ` · ${m365Count} profile${m365Count === 1 ? '' : 's'}` : ' · No data'}
          </ToggleButton>
          <ToggleButton
            icon={<LockClosed20Regular />}
            checked={!m365Open && paOpen && !iaOpen}
            disabled={!hasPaData}
            onClick={() => { setM365Open(false); setPaOpen(true); setIaOpen(false); }}
          >
            Private Access{hasPaData ? ` · ${paAppCount} app${paAppCount === 1 ? '' : 's'}` : ' · No data'}
          </ToggleButton>
          <ToggleButton
            icon={<Globe20Regular />}
            checked={!m365Open && !paOpen && iaOpen}
            onClick={() => { setM365Open(false); setIaOpen(true); setPaOpen(false); }}
          >
            Internet Access &middot; {groups.length} profile{groups.length === 1 ? '' : 's'}
          </ToggleButton>
        </div>

        {/* ── Microsoft 365 section ────────────────────────────────── */}
        {hasMsData && (
          <div className={mergeClasses(styles.sectionCard, m365Open && styles.sectionCardOpen)}>
            <div
              className={mergeClasses(
                styles.sectionHead,
                !m365Open && styles.sectionHeadCollapsed,
              )}
              role="button"
              tabIndex={0}
              onClick={() => setM365Open((v) => !v)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  setM365Open((v) => !v);
                }
              }}
            >
              <LockShield20Regular />
              <Text className={styles.sectionTitle}>Microsoft 365</Text>
              <span className={styles.sectionGrow} />
              <span className={styles.sectionChevron}>
                {m365Open ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
              </span>
            </div>
            {m365Open && (
              <div className={styles.iaSectionBody}>
                <ForwardingProfilesCard
                  profiles={data?.forwardingProfiles ?? []}
                  types={['m365']}
                />
              </div>
            )}
          </div>
        )}

        {/* ── Private Access section ────────────────────────────────── */}
        <div className={mergeClasses(styles.sectionCard, hasPaData && paOpen && styles.sectionCardOpen)}>
          <div
            className={mergeClasses(
              styles.sectionHead,
              (!paOpen || !hasPaData) && styles.sectionHeadCollapsed,
              !hasPaData && styles.sectionHeadDisabled,
            )}
            role="button"
            tabIndex={hasPaData ? 0 : -1}
            onClick={() => { if (hasPaData) setPaOpen((v) => !v); }}
            onKeyDown={(e) => {
              if (!hasPaData) return;
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                setPaOpen((v) => !v);
              }
            }}
          >
            <LockClosed20Regular />
            <Text className={styles.sectionTitle}>Private Access</Text>
            <span className={styles.sectionGrow} />
            {hasPaData && (
              <span className={styles.sectionChevron}>
                {paOpen ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
              </span>
            )}
          </div>
          {hasPaData && paOpen ? (
            <div className={styles.paSectionBody}>
              <PrivateAccessView
                privateAccess={data?.privateAccess ?? { apps: [], appProxyApps: [], authStrength: [] }}
                caDetails={data?.caDetails ?? []}
                forwardingProfiles={data?.forwardingProfiles ?? []}
              />
            </div>
          ) : !hasPaData ? (
            <div className={styles.paEmptyState}>
              <Text>No Private Access data in this data source.</Text>
              <Text size={200}>
                Load a data file exported with{' '}
                <code>Export-GsaFixture.ps1</code> to inspect PA apps and App
                Proxy applications.
              </Text>
            </div>
          ) : null}
        </div>

        {/* ── Internet Access section ───────────────────────────────── */}
        <div className={mergeClasses(styles.sectionCard, iaOpen && styles.sectionCardOpen)}>
          <div
            className={mergeClasses(
              styles.sectionHead,
              !iaOpen && styles.sectionHeadCollapsed,
            )}
            role="button"
            tabIndex={0}
            onClick={() => setIaOpen((v) => !v)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                setIaOpen((v) => !v);
              }
            }}
          >
            <Globe20Regular />
            <Text className={styles.sectionTitle}>Internet Access</Text>
            <span className={styles.sectionGrow} />
            <span className={styles.sectionChevron}>
              {iaOpen ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
            </span>
          </div>
          {iaOpen && (
            <div className={styles.iaSectionBody}>
              <ForwardingProfilesCard
                profiles={data?.forwardingProfiles ?? []}
                types={['internet']}
              />
              <TenantPostureStrip
                profiles={data?.profiles ?? []}
                tenantPolicies={data?.tenantPolicies ?? []}
              />
              <PolicyTable groups={groups} profiles={data?.profiles ?? []} />
            </div>
          )}
        </div>

      </div>
    </AppShell>
  );
}
