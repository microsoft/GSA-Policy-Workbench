/**
 * Browser-side fixture exporter — serializes the in-memory ProfileTreeResult
 * (already fully loaded by loader.ts) to the same JSON shape that
 * parseFixture() consumes, then triggers a file download.
 *
 * No Graph calls are made — all data is already in memory after loadProfileTree().
 * The output is safe to reload via the file picker on any workbench instance.
 */

import type { ProfileTreeResult } from '../graph/loader';

/**
 * Trigger a browser download of the current ProfileTreeResult as a fixture
 * JSON file.  The caller is responsible for supplying a meaningful filename
 * (e.g. "mountainsparadise.com_2026-07-31T14-30-00Z.json").
 */
export function downloadFixture(data: ProfileTreeResult, filename: string): void {
  const label = filename.replace(/\.json$/, '');
  // Build as a plain object — the shape matches FilteringProfilesBody exactly
  // but we add _comment (not in the interface) without a type cast.
  const body: Record<string, unknown> = {
    _comment:
      `Exported from ${label} via GSA Policy Workbench on ${new Date().toUTCString()}. ` +
      'Contains real tenant data — do not commit to version control.',
    '@odata.context':
      'https://graph.microsoft.com/beta/$metadata#networkAccess/filteringProfiles',
    value: data.profiles,
  };

  if (data.caDetails.length > 0)                  body['conditionalAccessPolicies']       = data.caDetails;
  if (data.directory.length > 0)                  body['directoryObjects']                = data.directory;
  if (data.privateAccess.authStrength.length > 0) body['authenticationStrengthPolicies']  = data.privateAccess.authStrength;
  if (data.privateAccess.apps.length > 0)         body['privateAccessApps']               = data.privateAccess.apps;
  if (data.privateAccess.appProxyApps.length > 0) body['appProxyApps']                    = data.privateAccess.appProxyApps;
  if (data.forwardingProfiles.length > 0)         body['forwardingProfiles']              = data.forwardingProfiles;
  if (data.tenantPolicies.length > 0)             body['tenantPolicies']                  = data.tenantPolicies;

  const json = JSON.stringify(body, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  try {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    // Append/remove required for Firefox — click outside the DOM is ignored.
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Build the fixture filename from a UPN (e.g. "user@contoso.com" → "contoso.com_2026-07-31T14-30-00Z.json"). */
export function fixtureFilename(upn: string): string {
  const domain = upn.includes('@') ? upn.split('@')[1] : upn;
  const ts = new Date().toISOString().replace(/:/g, '-').replace(/\.\d+Z$/, 'Z');
  return `${domain}_${ts}.json`;
}
