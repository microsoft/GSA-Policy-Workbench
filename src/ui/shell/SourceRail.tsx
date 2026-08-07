/**
 * SourceRail — collapsible left-side panel that lists registered data sources
 * and lets the user add, switch between, or remove them.
 *
 * Sources:
 *   • Live  — one authenticated Microsoft Graph connection (one max).
 *   • File  — any number of local fixture files loaded via the file picker.
 *
 * The rail collapses to a narrow icon strip (~44 px) to maximise content area.
 */

import { useRef, useState, type ChangeEvent } from 'react';
import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Text,
  Tooltip,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import {
  Add20Regular,
  ChevronLeft20Regular,
  ChevronRight20Regular,
  Dismiss16Regular,
  Document20Regular,
  Globe20Regular,
  ArrowEnter20Regular,
} from '@fluentui/react-icons';
import { useIsAuthenticated } from '@azure/msal-react';
import { useDataSource } from '../../app/dataSourceContext';
import { useAuth } from '../../auth/authContext';
import { friendlyError } from '../friendlyError';
import { DEFAULT_TENANT } from '../../auth/msalConfig';

// ── styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  rail: {
    display: 'flex',
    flexDirection: 'column',
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    flexShrink: 0,
    overflow: 'hidden',
    transition: 'width 0.18s ease',
    // Ensure children don't shrink the rail during transition
    minWidth: 0,
  },
  railOpen:   { width: '220px' },
  railClosed: { width: '44px' },

  // Header: "Sources" label + collapse chevron
  header: {
    flexShrink: 0,
    display: 'flex',
    alignItems: 'center',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalXS}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    gap: tokens.spacingHorizontalXS,
    minHeight: '40px',
  },
  headerTitle: {
    flex: 1,
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    paddingLeft: tokens.spacingHorizontalXS,
  },

  // Source list
  list: {
    flex: 1,
    overflowY: 'auto',
    overflowX: 'hidden',
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
  },

  // Individual source row
  item: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
    paddingLeft: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalXS,
    cursor: 'pointer',
    userSelect: 'none',
    minWidth: 0,
    borderLeft: `3px solid transparent`,
    ':hover': { backgroundColor: tokens.colorNeutralBackground2Hover },
  },
  itemActive: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderLeftColor: tokens.colorBrandStroke1,
    ':hover': { backgroundColor: tokens.colorNeutralBackground3Hover },
  },
  itemIcon: {
    flexShrink: 0,
    color: tokens.colorNeutralForeground2,
    display: 'inline-flex',
  },
  itemIconActive: { color: tokens.colorBrandForeground1 },
  itemLabel: {
    flex: 1,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    minWidth: 0,
  },

  // Footer: Add source button
  footer: {
    flexShrink: 0,
    padding: tokens.spacingHorizontalS,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },

  error: {
    padding: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalS}`,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorPaletteRedForeground1,
    wordBreak: 'break-word',
  },
});

// ── component ────────────────────────────────────────────────────────────────

export function SourceRail() {
  const styles = useStyles();
  const [open, setOpen] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const {
    sources, activeSourceId, sourceError,
    activateSource, removeSource, addFileSource, addLiveSource,
  } = useDataSource();
  const { signIn, account } = useAuth();
  const isAuthenticated = useIsAuthenticated();

  // ── file picker ───────────────────────────────────────────────────────────

  async function handleFileChosen(e: ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setError(null);
    try {
      const text = await file.text();
      addFileSource(text, file.name);
    } catch (err) {
      // The file name is deliberately not echoed: exported fixture names carry
      // the tenant domain by construction (finding 15).
      setError(friendlyError(err, 'That file could not be loaded.'));
    }
  }

  // ── add live ──────────────────────────────────────────────────────────────

  async function handleAddLive() {
    setError(null);
    if (isAuthenticated && account) {
      // Already signed in — just register the entry.
      addLiveSource(account.username);
      return;
    }
    // Not signed in — trigger MSAL redirect; the live entry is added in App.tsx
    // once useIsAuthenticated() flips to true.
    try {
      await signIn(DEFAULT_TENANT);
    } catch (err) {
      setError(friendlyError(err, 'Sign-in failed.'));
    }
  }

  // ── derived ───────────────────────────────────────────────────────────────

  const hasLive = sources.some(s => s.type === 'live');

  // ── render ────────────────────────────────────────────────────────────────

  return (
    <aside
      className={mergeClasses(styles.rail, open ? styles.railOpen : styles.railClosed)}
      aria-label="Data sources"
    >
      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".json"
        style={{ display: 'none' }}
        onChange={handleFileChosen}
      />

      {/* ── Header ────────────────────────────────────────────────────── */}
      <div className={styles.header}>
        {open && <span className={styles.headerTitle}>Sources</span>}
        <Tooltip
          content={open ? 'Collapse panel' : 'Expand panel'}
          relationship="label"
        >
          <Button
            appearance="subtle"
            size="small"
            icon={open ? <ChevronLeft20Regular /> : <ChevronRight20Regular />}
            onClick={() => setOpen(v => !v)}
            aria-label={open ? 'Collapse sources panel' : 'Expand sources panel'}
          />
        </Tooltip>
      </div>

      {/* ── Source list ───────────────────────────────────────────────── */}
      <div className={styles.list} role="listbox" aria-label="Data sources">
        {sources.map(src => {
          const isActive = src.id === activeSourceId;
          const Icon = src.type === 'live' ? Globe20Regular : Document20Regular;
          const row = (
            <div
              key={src.id}
              role="option"
              aria-selected={isActive}
              className={mergeClasses(styles.item, isActive && styles.itemActive)}
              onClick={() => activateSource(src.id)}
              onKeyDown={e => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  activateSource(src.id);
                }
              }}
              tabIndex={0}
            >
              <span className={mergeClasses(styles.itemIcon, isActive && styles.itemIconActive)}>
                <Icon />
              </span>
              {open && (
                <>
                  <Text className={styles.itemLabel} title={src.label}>
                    {src.label}
                  </Text>
                  <Tooltip content="Remove source" relationship="label">
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<Dismiss16Regular />}
                      aria-label={`Remove ${src.label}`}
                      onClick={e => {
                        e.stopPropagation();
                        removeSource(src.id);
                      }}
                    />
                  </Tooltip>
                </>
              )}
            </div>
          );

          // Collapsed: wrap each icon in a tooltip showing the label
          return open ? row : (
            <Tooltip key={src.id} content={src.label} relationship="label">
              {row}
            </Tooltip>
          );
        })}
      </div>

      {/* ── Error ─────────────────────────────────────────────────────── */}
      {(error ?? sourceError) && open && (
        <Text className={styles.error}>{error ?? sourceError}</Text>
      )}

      {/* ── Footer: Add source ────────────────────────────────────────── */}
      <div className={styles.footer}>
        {open ? (
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button
                appearance="outline"
                size="small"
                icon={<Add20Regular />}
                style={{ width: '100%' }}
              >
                Add source
              </Button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem
                  icon={<ArrowEnter20Regular />}
                  disabled={hasLive}
                  onClick={() => void handleAddLive()}
                >
                  {hasLive ? 'Live already added' : 'Connect live (Microsoft Entra)'}
                </MenuItem>
                <MenuItem
                  icon={<Document20Regular />}
                  onClick={() => fileInputRef.current?.click()}
                >
                  Load fixture file
                </MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        ) : (
          <Tooltip content="Add source" relationship="label">
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<Add20Regular />}
                  aria-label="Add source"
                />
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem
                    icon={<ArrowEnter20Regular />}
                    disabled={hasLive}
                    onClick={() => void handleAddLive()}
                  >
                    {hasLive ? 'Live already added' : 'Connect live (Microsoft Entra)'}
                  </MenuItem>
                  <MenuItem
                    icon={<Document20Regular />}
                    onClick={() => fileInputRef.current?.click()}
                  >
                    Load fixture file
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </Tooltip>
        )}
      </div>
    </aside>
  );
}
