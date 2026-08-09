/**
 * Adaptive Security Profile tree loader — Tier 3.
 *
 * Loads the full Security Profile → policyLink → filteringPolicy → rules tree
 * from Graph using an adaptive two-branch strategy:
 *
 *   Branch A (best case): one deep $expand call returns policyRules inline.
 *   Branch B (fallback):  rules absent — fetch per unique policy via $batch.
 *
 * This loader sits in Tier 3 (Graph Connector layer) and uses the generic
 * connector. It never talks to Graph directly — all calls go through connector
 * or graphBatch from client.ts.
 */

import { graphGet, graphBatch, graphPost, GraphError } from './client';
import { recordGraphCall } from './interceptor';
import { REQUIRED_SCOPES, CA_DETAIL_SCOPE, DIRECTORY_SCOPE, APPLICATION_SCOPE, USER_READ_SCOPE, GROUP_MEMBER_SCOPE } from '../../auth/scopes';
import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type {
  FilteringPolicyLink,
  WebContentFilteringPolicy,
} from '../../model/definitions/WebContentFilteringPolicy.definition';
import type { WebCategory } from '../../model/definitions/WebContentFilteringPolicy.definition';
import type {
  ConditionalAccessPolicy,
  DirectoryObjectRef,
} from '../../model/definitions/ConditionalAccessPolicy.definition';
import type {
  PrivateAccessApp,
  AppProxyApp,
  AuthenticationStrengthPolicy,
  ApplicationSegment,
} from '../../model/definitions/PrivateAccessApp.definition';
import type { ForwardingProfile, ForwardingRule } from '../../model/definitions/ForwardingProfile.definition';
import type { ODataCollection } from '../../model/types';

// ---------------------------------------------------------------------------
// Result type
// ---------------------------------------------------------------------------

export type LoadBranch = 'A-inline' | 'B-batched';

/**
 * The `filteringProfiles` response body. Fixtures (and only fixtures) carry
 * full Conditional Access policy detail as a top-level sibling of `value`;
 * live Graph omits it (the SPA holds only CA stubs on each profile).
 */
interface ProfilesBody extends ODataCollection<SecurityProfile> {
  conditionalAccessPolicies?: ConditionalAccessPolicy[];
  directoryObjects?: DirectoryObjectRef[];
  authenticationStrengthPolicies?: AuthenticationStrengthPolicy[];
  privateAccessApps?: PrivateAccessApp[];
  appProxyApps?: AppProxyApp[];
  forwardingProfiles?: ForwardingProfile[];
  tenantPolicies?: TenantPolicy[];
}

/**
 * A tenant-wide policy as listed by `/networkAccess/{collection}` — the minimal
 * projection (id + name + type) needed to cross-reference against the profile
 * tree for orphaned-policy detection and tenant-limit counters (spec §10, V2
 * items 4 & 6). Read-only inspection signal — uses `NetworkAccess.Read.All`
 * (already required), so it loads live as well as from a fixture sibling.
 */
export interface TenantPolicy {
  id: string;
  name: string;
  /** Concrete policy `@odata.type` (resolved to a friendly label in the UI). */
  '@odata.type'?: string;
}

export interface ProfileTreeResult {
  profiles: SecurityProfile[];
  /**
   * Full CA policy detail (conditions + grant controls) when available — used
   * to show user/group targeting. Empty when only stubs are available.
   */
  caDetails: ConditionalAccessPolicy[];
  /**
   * Resolved user/group display names for CA targeting (exporter only). Empty
   * in live Graph mode — the UI then falls back to raw object ids.
   */
  directory: DirectoryObjectRef[];
  /**
   * Private Access policy domain (V1+, read-only). Populated from a fixture's
   * top-level sibling arrays (spec §6); empty in live Graph mode — the V0 SPA
   * does not request the extra scopes those endpoints need (spec §6.5).
   */
  privateAccess: PrivateAccessDomain;
  /**
   * GSA traffic-forwarding profiles (internet / m365 / private) with their
   * enablement state — the "is GSA on for this traffic class?" signal
   * (spec §6.6). Uses `NetworkAccess.Read.All` (already required), so it loads
   * live as well as from a fixture.
   */
  forwardingProfiles: ForwardingProfile[];
  /**
   * Tenant-wide policy inventory (filtering / TLS / prompt / content) — the
   * minimal id + name + type projection used for orphaned-policy detection and
   * tenant-limit counters (spec §10, V2 items 4 & 6). Cross-referenced against
   * `profiles` in the UI. Uses `NetworkAccess.Read.All` (already required), so
   * it loads live as well as from a fixture sibling.
   */
  tenantPolicies: TenantPolicy[];
  branch: LoadBranch;
  /** Human-readable note surfaced in the UI (branch badge tooltip). */
  branchNote: string;
}

/** The Private Access domain bundle surfaced read-only from a fixture. */
export interface PrivateAccessDomain {
  apps: PrivateAccessApp[];
  appProxyApps: AppProxyApp[];
  authStrength: AuthenticationStrengthPolicy[];
}

const EMPTY_PRIVATE_ACCESS: PrivateAccessDomain = {
  apps: [],
  appProxyApps: [],
  authStrength: [],
};

// ---------------------------------------------------------------------------
// URL templates — recorded by the interceptor (no populated IDs)
// ---------------------------------------------------------------------------

const ENDPOINT_DEEP   = '/networkAccess/filteringProfiles?$expand=policies(policy(policyRules)),conditionalAccessPolicies';
const ENDPOINT_SHALLOW = '/networkAccess/filteringProfiles?$expand=policies(policy),conditionalAccessPolicies';
const ENDPOINT_RULES   = '/networkAccess/filteringPolicies/{id}/policyRules';

const DEEP_EXPAND_URL =
  '/networkAccess/filteringProfiles' +
  '?$expand=policies($expand=policy($expand=policyRules)),conditionalAccessPolicies';

const SHALLOW_EXPAND_URL =
  '/networkAccess/filteringProfiles' +
  '?$expand=policies($expand=policy),conditionalAccessPolicies';

const SCOPES = [...REQUIRED_SCOPES];
const FEATURE = 'load-table';

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Load the complete Security Profile tree.
 *
 * Tries Branch A (deep $expand returning rules inline). Falls back to
 * Branch B (shallow expand + batched policyRules per unique policy) if the
 * Graph beta API rejects the nested expand or omits the rules.
 */
export async function loadProfileTree(
  signal?: AbortSignal,
): Promise<ProfileTreeResult> {
  let profiles: SecurityProfile[];
  let caDetails: ConditionalAccessPolicy[] = [];
  let directory: DirectoryObjectRef[] = [];
  let privateAccess: PrivateAccessDomain = EMPTY_PRIVATE_ACCESS;
  let forwardingProfiles: ForwardingProfile[] = [];
  let tenantPolicies: TenantPolicy[] = [];
  let deepExpandWorked = false;

  function captureSiblings(res: ProfilesBody): void {
    if (res.privateAccessApps || res.appProxyApps || res.authenticationStrengthPolicies) {
      privateAccess = {
        apps: res.privateAccessApps ?? [],
        appProxyApps: res.appProxyApps ?? [],
        authStrength: res.authenticationStrengthPolicies ?? [],
      };
    }
    if (res.forwardingProfiles) {
      forwardingProfiles = res.forwardingProfiles;
    }
    if (res.tenantPolicies) {
      tenantPolicies = res.tenantPolicies;
    }
  }

  try {
    const res = await graphGet<ProfilesBody>(
      DEEP_EXPAND_URL,
      {
        scopes: SCOPES,
        feature: FEATURE,
        endpoint: ENDPOINT_DEEP,
        paginate: true,
        signal,
      },
    );
    profiles = res.value;
    caDetails = res.conditionalAccessPolicies ?? [];
    directory = res.directoryObjects ?? [];
    captureSiblings(res);
    deepExpandWorked = true;
  } catch (err) {
    if (err instanceof GraphError && err.status >= 400 && err.status < 500) {
      // Deep nested $expand rejected by the beta API — fall back to shallow.
      const res = await graphGet<ProfilesBody>(
        SHALLOW_EXPAND_URL,
        {
          scopes: SCOPES,
          feature: FEATURE,
          endpoint: ENDPOINT_SHALLOW,
          paginate: true,
          signal,
        },
      );
      profiles = res.value;
      caDetails = res.conditionalAccessPolicies ?? [];
      directory = res.directoryObjects ?? [];
      captureSiblings(res);

      recordGraphCall({
        ts: new Date().toISOString(),
        method: 'GET',
        endpoint: ENDPOINT_DEEP,
        scopes: SCOPES,
        status: err.status,
        durationMs: 0,
        apiVersion: 'beta',
        feature: FEATURE,
        fallback: {
          reason: `3-level $expand (policyRules) rejected with ${err.status}`,
          alternativePath: ENDPOINT_RULES,
        },
        notes: 'Branch B selected: deep $expand not supported.',
      });
    } else {
      throw err;
    }
  }

  // Live Graph does not return CA detail as a sibling — only stubs sit on each
  // profile. When no detail arrived (i.e. not a fixture) fetch it live, best
  // effort, using the optional Policy.Read.All scope. Missing scope → skip
  // silently so the table still renders (spec §2.4).
  if (caDetails.length === 0) {
    caDetails = await fetchLinkedCaDetails(profiles, signal);
  }

  // Resolve the user/group object ids targeted by those CA policies to friendly
  // display names (optional Directory.Read.All — degrades to raw GUIDs when the
  // scope is not granted). Only needed live; a fixture already carries resolved
  // names in its `directoryObjects` sibling.
  if (directory.length === 0 && caDetails.length > 0) {
    directory = await fetchDirectoryObjects(caDetails, signal);
  }

  // Forwarding-profile state uses NetworkAccess.Read.All (already required), so
  // fetch it live when a fixture did not supply it (spec §6.6). Best-effort.
  if (forwardingProfiles.length === 0) {
    forwardingProfiles = await fetchForwardingProfiles(signal);
  }

  // Tenant-wide policy inventory (orphaned-policy detection + tenant-limit
  // counters, spec §10 V2 items 4 & 6). Same already-required scope, so fetch
  // it live when a fixture did not supply it. Best-effort.
  if (tenantPolicies.length === 0) {
    tenantPolicies = await fetchTenantPolicies(signal);
  }

  // Live Private Access domain (Directory.Read.All + Application.Read.All —
  // both optional). Only called when a fixture did not supply PA data.
  // In file mode the fileTransport returns 404 for these endpoints, which is
  // caught inside fetchPrivateAccessDomain and degrades to empty arrays.
  if (privateAccess === EMPTY_PRIVATE_ACCESS) {
    privateAccess = await fetchPrivateAccessDomain(signal);
  }

  // Did policyRules come back inline? If so → Branch A.
  const rulesInline = deepExpandWorked && anyPolicyHasInlineRules(profiles);

  if (rulesInline) {
    return {
      profiles,
      caDetails,
      directory,
      privateAccess,
      forwardingProfiles,
      tenantPolicies,
      branch: 'A-inline',
      branchNote:
        'Deep $expand returned policyRules inline — entire tree in one call.',
    };
  }

  // Branch B: fetch rules per unique policy via $batch and attach.
  await attachRulesByBatch(profiles, signal);

  return {
    profiles,
    caDetails,
    directory,
    privateAccess,
    forwardingProfiles,
    tenantPolicies,
    branch: 'B-batched',
    branchNote: deepExpandWorked
      ? 'Deep $expand succeeded but omitted policyRules — fetched per-policy via $batch.'
      : 'Deep $expand unsupported — used shallow expand + per-policy $batch.',
  };
}

// ---------------------------------------------------------------------------
// Live Conditional Access detail (optional — Policy.Read.All)
// ---------------------------------------------------------------------------

/**
 * Fetch full Conditional Access detail for every CA policy linked to a profile.
 *
 * Live Graph projects only stubs (`id` + `displayName`) onto each profile, so
 * we batch-GET each linked policy by id to obtain conditions + grant controls.
 * Requires the optional `Policy.Read.All` scope; if it is not granted the
 * token provider returns null and the call fails — we swallow that and return
 * an empty array so the table still loads (spec §2.4, graceful degradation).
 */
async function fetchLinkedCaDetails(
  profiles: SecurityProfile[],
  signal?: AbortSignal,
): Promise<ConditionalAccessPolicy[]> {
  const ids = new Set<string>();
  for (const profile of profiles) {
    for (const stub of profile.conditionalAccessPolicies ?? []) {
      if (stub?.id) ids.add(stub.id);
    }
  }
  if (ids.size === 0) return [];

  const uniqueIds = [...ids];
  const requests = uniqueIds.map((id) => ({
    id,
    method: 'GET' as const,
    url: `/identity/conditionalAccess/policies/${id}`,
  }));

  try {
    const responses = await graphBatch<ConditionalAccessPolicy>(requests, {
      scopes: [CA_DETAIL_SCOPE],
      feature: FEATURE,
    });

    const details: ConditionalAccessPolicy[] = [];
    for (const id of uniqueIds) {
      const resp = responses.get(id);
      if (resp && resp.status >= 200 && resp.status < 300) {
        details.push(resp.body);
      }
    }
    return details;
  } catch {
    // No Policy.Read.All consent (or a transient failure) — degrade to stubs.
    return [];
  } finally {
    void signal;
  }
}

// ---------------------------------------------------------------------------
// Live directory name resolution (optional — Directory.Read.All)
// ---------------------------------------------------------------------------

const GUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const ENDPOINT_GET_BY_IDS = '/directoryObjects/getByIds';

/**
 * Resolve the user/group object ids referenced by CA user conditions to
 * friendly display names via `POST /directoryObjects/getByIds`.
 *
 * Live Graph projects only raw GUIDs into the CA conditions; this turns them
 * into `displayName` so the where-used targeting reads names instead of ids
 * (matching what the exporter bakes into a fixture's `directoryObjects`).
 * Requires the optional `Directory.Read.All` scope — when it is not granted the
 * token provider returns null and the POST fails, which we swallow so the table
 * still loads with raw GUIDs (graceful degradation). Only user + group ids are
 * resolved; role-template and special tokens (All / GuestsOrExternalUsers) are
 * skipped. The ids travel in the request **body**, so the audit endpoint stays
 * templated.
 */
async function fetchDirectoryObjects(
  caDetails: ConditionalAccessPolicy[],
  signal?: AbortSignal,
): Promise<DirectoryObjectRef[]> {
  const ids = new Set<string>();
  for (const ca of caDetails) {
    const users = ca.conditions?.users;
    if (!users) continue;
    for (const list of [
      users.includeUsers,
      users.excludeUsers,
      users.includeGroups,
      users.excludeGroups,
    ]) {
      for (const id of list ?? []) {
        if (GUID_RE.test(id)) ids.add(id);
      }
    }
  }
  if (ids.size === 0) return [];

  try {
    const res = await graphPost<ODataCollection<DirectoryObjectRef>>(
      ENDPOINT_GET_BY_IDS,
      { ids: [...ids], types: ['user', 'group'] },
      {
        scopes: [DIRECTORY_SCOPE],
        feature: FEATURE,
        endpoint: ENDPOINT_GET_BY_IDS,
        signal,
      },
    );
    return res.value ?? [];
  } catch {
    // No Directory.Read.All consent (or a transient failure) — keep raw GUIDs.
    return [];
  }
}

// ---------------------------------------------------------------------------
// Forwarding-profile state (NetworkAccess.Read.All — already required)
// ---------------------------------------------------------------------------

const ENDPOINT_FORWARDING =
  '/networkAccess/forwardingProfiles' +
  '?$select=id,name,description,version,lastModifiedDateTime,priority,state,trafficForwardingType,associations' +
  '&$expand=policies($expand=policy)';

/**
 * Fetch the GSA traffic-forwarding profiles (internet / m365 / private) —
 * state, priority, associations, and their linked acquisition rules. Uses the
 * already-required `NetworkAccess.Read.All` scope, so
 * no new consent is needed. Best-effort: a transient failure degrades to an
 * empty list and the status strip simply shows nothing rather than blocking
 * the view.
 */
async function fetchForwardingProfiles(
  signal?: AbortSignal,
): Promise<ForwardingProfile[]> {
  let profiles: ForwardingProfile[];
  try {
    const res = await graphGet<ODataCollection<ForwardingProfile>>(
      ENDPOINT_FORWARDING,
      {
        scopes: SCOPES,
        feature: FEATURE,
        endpoint: ENDPOINT_FORWARDING,
        paginate: true,
        signal,
      },
    );
    profiles = res.value ?? [];
  } catch {
    return [];
  }
  if (profiles.length > 0) {
    await attachForwardingPolicyRules(profiles, signal);
  }
  return profiles;
}

/**
 * The beta `$expand` only nests 2 levels (`policies($expand=policy)`) —
 * a 3rd level for `policyRules` is rejected, same limitation as the main
 * profile tree (see the Branch A/B comment above). So the acquisition rules
 * for every linked forwarding policy are fetched in one `$batch` call and
 * attached back onto `policy.policyRules` in place. Best-effort: a failed
 * batch leaves profiles without rule detail rather than failing the load.
 */
async function attachForwardingPolicyRules(
  profiles: ForwardingProfile[],
  signal?: AbortSignal,
): Promise<void> {
  const targets = new Map<string, { policyRules?: ForwardingRule[] }>();
  for (const profile of profiles) {
    for (const link of profile.policies ?? []) {
      if (link.policy?.id) targets.set(link.policy.id, link.policy);
    }
  }

  const uniqueIds = [...targets.keys()];
  if (uniqueIds.length === 0) return;

  const requests = uniqueIds.map((id) => ({
    id,
    method: 'GET' as const,
    url: `/networkAccess/forwardingPolicies/${id}/policyRules`,
  }));

  try {
    const responses = await graphBatch<ODataCollection<ForwardingRule>>(
      requests,
      { scopes: SCOPES, feature: FEATURE },
    );
    for (const [id, policy] of targets) {
      const resp = responses.get(id);
      policy.policyRules =
        resp && resp.status >= 200 && resp.status < 300
          ? (resp.body.value ?? [])
          : [];
    }
  } catch {
    // No consent or a transient failure — profiles keep showing state only.
  }
  void signal;
}

// ---------------------------------------------------------------------------
// Tenant-wide policy inventory (NetworkAccess.Read.All — already required)
// ---------------------------------------------------------------------------

/**
 * The tenant-wide policy collections, with the concrete `@odata.type` each one
 * holds. `$select=id,name` keeps the projection minimal — we only need to
 * count policies and cross-reference their ids against the profile tree
 * (orphaned-policy detection + tenant-limit counters, spec §10 V2 items 4 & 6).
 */
const TENANT_POLICY_COLLECTIONS: ReadonlyArray<{ collection: string; odataType: string }> = [
  { collection: 'filteringPolicies', odataType: '#microsoft.graph.networkaccess.webContentFilteringPolicy' },
  { collection: 'tlsInspectionPolicies', odataType: '#microsoft.graph.networkaccess.tlsInspectionPolicy' },
  { collection: 'promptPolicies', odataType: '#microsoft.graph.networkaccess.promptPolicy' },
  { collection: 'filePolicies', odataType: '#microsoft.graph.networkaccess.filePolicy' },
  { collection: 'cloudFirewallPolicies', odataType: '#microsoft.graph.networkaccess.cloudFirewallPolicy' },
];

interface TenantPolicyWire {
  id: string;
  name?: string;
  '@odata.type'?: string;
}

/**
 * List every policy in the tenant (across the filtering / TLS / prompt /
 * content collections). Read-only, best-effort per collection: a transient
 * failure on one collection degrades to skipping it rather than failing the
 * whole load. Uses the already-required `NetworkAccess.Read.All` scope.
 */
async function fetchTenantPolicies(signal?: AbortSignal): Promise<TenantPolicy[]> {
  const all: TenantPolicy[] = [];
  for (const { collection, odataType } of TENANT_POLICY_COLLECTIONS) {
    const endpoint = `/networkAccess/${collection}?$select=id,name`;
    try {
      const res = await graphGet<ODataCollection<TenantPolicyWire>>(endpoint, {
        scopes: SCOPES,
        feature: FEATURE,
        endpoint,
        paginate: true,
        signal,
      });
      for (const p of res.value ?? []) {
        if (!p?.id) continue;
        all.push({
          id: p.id,
          name: p.name ?? '',
          '@odata.type': p['@odata.type'] ?? odataType,
        });
      }
    } catch {
      // Collection unavailable (e.g. a beta gap) — skip it, keep the rest.
    }
  }
  return all;
}

// ---------------------------------------------------------------------------
// Private Access live fetch (Directory.Read.All + Application.Read.All — optional)
// ---------------------------------------------------------------------------

const PA_FEATURE = 'load-private-access';

/** OData filter selecting both PA app kinds (needs ConsistencyLevel: eventual). */
const PA_TAGS_FILTER =
  "(tags/any(t:t eq 'PrivateAccessNonWebApplication')" +
  " or tags/any(t:t eq 'NetworkAccessQuickAccessApplication'))";

const ADVANCED_QUERY_HEADERS = { ConsistencyLevel: 'eventual' } as const;

const ENDPOINT_PA_APPS        = '/applications (PA tag filter)';
const ENDPOINT_PA_SPS         = '/servicePrincipals (PA tag filter)';
const ENDPOINT_APPPROXY_LIST   = '/applications (App Proxy filter)';

/** Wire shape returned by `/applications/{id}?$select=id,onPremisesPublishing`. */
interface OnPremPublishingWire {
  isDnsResolutionEnabled?: boolean;
}

/**
 * Orchestrate the live Private Access fetch — SPs + App Proxy apps.
 * Both fetches are best-effort and degrade to empty on missing scopes.
 */
async function fetchPrivateAccessDomain(
  signal?: AbortSignal,
): Promise<PrivateAccessDomain> {
  const [apps, appProxyApps] = await Promise.all([
    fetchLivePrivateAccessApps(signal),
    fetchLiveAppProxyApps(signal),
  ]);
  return { apps, appProxyApps, authStrength: [] };
}

/**
 * Fetch tagged Private Access / Quick Access service principals and enrich them
 * with segment, DNS, and assignment data.
 *
 * Step 1 — SP list (v1.0, `Directory.Read.All`, advanced query).
 * Approach mirrors Migrate2GSA: application objects are the primary source
 * (tags + applicationObjectId come from /beta/applications). Service principals
 * are fetched in parallel for appRoleAssignmentRequired + assignments.
 *
 * All steps after step 1 are independently best-effort.
 */
async function fetchLivePrivateAccessApps(
  signal?: AbortSignal,
): Promise<PrivateAccessApp[]> {
  const SEGMENT_PATH =
    '/onPremisesPublishing/segmentsConfiguration' +
    '/microsoft.graph.ipSegmentConfiguration/applicationSegments';

  // Step 1: list PA application objects (Application.Read.All, beta).
  //
  // Migrate2GSA confirmed: PA tags ARE on /beta/applications, not just on SPs.
  // This directly gives the application object ID needed for the segments
  // endpoint — no separate ID-resolution step required.
  type RawApp = { id: string; appId?: string; displayName?: string; tags?: string[] };
  let rawApps: RawApp[];
  try {
    const res = await graphGet<ODataCollection<RawApp>>(
      `/applications?$filter=${PA_TAGS_FILTER}&$count=true` +
      `&$select=id,appId,displayName,tags`,
      {
        scopes: [APPLICATION_SCOPE],
        feature: PA_FEATURE,
        endpoint: ENDPOINT_PA_APPS,
        paginate: true,
        signal,
        extraHeaders: ADVANCED_QUERY_HEADERS,
      },
    );
    rawApps = res.value ?? [];
  } catch {
    // Application.Read.All not granted.
    return [];
  }
  if (rawApps.length === 0) return [];

  // Seed the PrivateAccessApp list — id starts as applicationObjectId;
  // it is replaced by the SP object id in step 3 when Directory.Read.All is available.
  const sps: PrivateAccessApp[] = rawApps.map(app => ({
    id: app.id,
    appId: app.appId,
    displayName: app.displayName,
    tags: app.tags,
    applicationObjectId: app.id,
  }));

  // Steps 2 + 3 run in parallel — independent scopes, no data dependency.
  await Promise.all([

    // Step 2: segments + DNS flag — one GET per app (mirrors Migrate2GSA).
    //
    // Direct graphGet calls avoid the /beta/$batch URL-resolution ambiguity;
    // Migrate2GSA confirms this is the correct pattern for these sub-paths.
    (async () => {
      try {
        await Promise.all(sps.map(async sp => {
          if (!sp.applicationObjectId) return;
          const appObjId = sp.applicationObjectId;

          // 2a — isDnsResolutionEnabled from onPremisesPublishing
          try {
            const pub = await graphGet<{ id?: string; onPremisesPublishing?: OnPremPublishingWire }>(
              `/applications/${appObjId}?$select=id,onPremisesPublishing`,
              {
                scopes: [APPLICATION_SCOPE],
                feature: PA_FEATURE,
                endpoint: '/applications/{id}',
              },
            );
            sp.isDnsResolutionEnabled = pub.onPremisesPublishing?.isDnsResolutionEnabled;
          } catch { /* best-effort */ }

          // 2b — applicationSegments (navigation property, must use its own sub-path)
          try {
            const seg = await graphGet<ODataCollection<ApplicationSegment>>(
              `/applications/${appObjId}${SEGMENT_PATH}`,
              {
                scopes: [APPLICATION_SCOPE],
                feature: PA_FEATURE,
                endpoint: `/applications/{id}${SEGMENT_PATH}`,
              },
            );
            sp.applicationSegments = seg.value ?? [];
          } catch { /* best-effort */ }
        }));
      } catch (err) {
        if (import.meta.env.DEV) console.error('[PA] step2 ERROR:', err);
      }
    })(),

    // Step 3: SP metadata (Directory.Read.All) — appRoleAssignmentRequired + assignments.
    // Best-effort: absence degrades the assignment column only.
    (async () => {
      try {
        type RawSp = { id: string; appId?: string; appRoleAssignmentRequired?: boolean };
        const spRes = await graphGet<ODataCollection<RawSp>>(
          `/servicePrincipals?$filter=${PA_TAGS_FILTER}&$count=true` +
          `&$select=id,appId,appRoleAssignmentRequired`,
          {
            scopes: [DIRECTORY_SCOPE],
            feature: PA_FEATURE,
            endpoint: ENDPOINT_PA_SPS,
            paginate: true,
            signal,
            apiVersion: 'v1.0',
            extraHeaders: ADVANCED_QUERY_HEADERS,
          },
        );
        const spByAppId = new Map(
          (spRes.value ?? []).filter(s => s.appId).map(s => [s.appId!, s]),
        );
        // Overwrite id with SP object id (needed for /appRoleAssignedTo calls).
        for (const sp of sps) {
          const match = sp.appId ? spByAppId.get(sp.appId) : undefined;
          if (match) {
            sp.id = match.id;
            sp.appRoleAssignmentRequired = match.appRoleAssignmentRequired;
          }
        }
        // Direct GET per SP for assignments — appRoleAssignedTo is a navigation
        // property; $batch is unreliable for nav properties (same issue as segments).
        const spsWithSpId = sps.filter(sp => sp.id !== sp.applicationObjectId);
        await Promise.all(spsWithSpId.map(async sp => {
          try {
            const res = await graphGet<ODataCollection<{
              principalId?: string;
              principalType?: string;
              principalDisplayName?: string;
            }>>(
              `/servicePrincipals/${sp.id}/appRoleAssignedTo` +
              `?$select=principalId,principalType,principalDisplayName`,
              {
                scopes: [DIRECTORY_SCOPE],
                feature: PA_FEATURE,
                endpoint: '/servicePrincipals/{id}/appRoleAssignedTo',
                apiVersion: 'v1.0',
              },
            );
            sp.appRoleAssignedTo = res.value ?? [];
          } catch { /* best-effort per SP */ }
        }));
      } catch { /* Directory.Read.All not granted — assignment data absent */ }
    })(),

  ]);

  return sps;
}

/**
 * Fetch Application Proxy applications using the mandatory two-call pattern.
 *
 * The bulk list endpoint returns a wrong default (`passthru`) for
 * `externalAuthenticationType` regardless of the real value (spec §6.2).
 * The per-app call returns the accurate value. Requires `Application.Read.All`.
 * Best-effort — returns [] on missing scope or transient failure.
 */
async function fetchLiveAppProxyApps(signal?: AbortSignal): Promise<AppProxyApp[]> {
  // Step 1: list App Proxy apps (accurate externalAuthenticationType is NOT here).
  let appList: { id: string; displayName?: string }[];
  try {
    const res = await graphGet<ODataCollection<{ id: string; displayName?: string }>>(
      '/applications?$filter=onPremisesPublishing/isOnPremPublishingEnabled eq true' +
      '&$select=id,displayName',
      {
        scopes: [APPLICATION_SCOPE],
        feature: PA_FEATURE,
        endpoint: ENDPOINT_APPPROXY_LIST,
        paginate: true,
        signal,
      },
    );
    appList = res.value ?? [];
  } catch {
    return []; // Application.Read.All not granted.
  }
  if (appList.length === 0) return [];

  // Step 2: per-app accurate onPremisesPublishing (correct externalAuthenticationType).
  const requests = appList.map(app => ({
    id: app.id,
    method: 'GET' as const,
    url: `/beta/applications/${app.id}?$select=id,displayName,onPremisesPublishing`,
  }));
  try {
    const results = await graphBatch<AppProxyApp>(requests, {
      scopes: [APPLICATION_SCOPE],
      feature: PA_FEATURE,
    });
    const apps: AppProxyApp[] = [];
    for (const [, r] of results) {
      if (r.status === 200) apps.push(r.body);
    }
    return apps;
  } catch {
    // Batch failed — fall back to list data (without accurate externalAuthType).
    return appList.map(app => ({ id: app.id, displayName: app.displayName }));
  }
}

// ---------------------------------------------------------------------------
// Web content category lookup (NetworkAccess.Read.All — already required)
// ---------------------------------------------------------------------------

/**
 * Web content category lookup (Tier 3) — resolve which content category a URL
 * or host belongs to via the beta `getWebCategoriesByUrl` function (preview).
 *
 * Live-only: there is no fixture for arbitrary URL lookups, so this is not part
 * of the file-mode data source. Uses the already-required `NetworkAccess.Read.All`
 * scope (no new consent). The audit endpoint is templated — the queried host is
 * passed as the `@url` parameter and is never written to the call log.
 */
export async function fetchWebCategory(
  url: string,
  signal?: AbortSignal,
): Promise<WebCategory> {
  const fn =
    "/networkAccess/connectivity/microsoft.graph.networkaccess.getWebCategoriesByUrl(url='@url')";
  // The beta endpoint may return either a single entity or an OData collection
  // depending on the tenant or API version — unwrap either shape.
  const raw = await graphGet<WebCategory & { value?: WebCategory[] }>(
    `${fn}?@url=${encodeURIComponent(url)}`,
    {
      scopes: SCOPES,
      feature: 'web-category-lookup',
      endpoint: '/networkAccess/connectivity/getWebCategoriesByUrl',
      signal,
    },
  );
  if (Array.isArray(raw.value)) {
    const cat = raw.value[0];
    if (!cat?.name && !cat?.displayName) throw new Error(`No web content category found for '${url}'`);
    return cat;
  }
  return raw;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function anyPolicyHasInlineRules(profiles: SecurityProfile[]): boolean {
  return profiles.some((p) =>
    (p.policies ?? []).some(
      (link) =>
        (link.policy as WebContentFilteringPolicy | undefined)?.policyRules !==
        undefined,
    ),
  );
}

/**
 * Collect every unique policy referenced across all profiles, batch-fetch its
 * rules from the endpoint that matches the policy kind (max 20 per $batch
 * call), and attach them back onto the policy objects in-place.
 *
 * Different policy kinds expose rules on different sub-routes — web content
 * filtering on `/filteringPolicies/{id}/policyRules`, TLS inspection on
 * `/tlsInspectionPolicies/{id}/policyRules`, and prompt protection on
 * `/promptPolicies/{id}/policyRules`. Routing every policy through the
 * filtering route (the old behaviour) silently dropped TLS and prompt rules.
 * All three use `NetworkAccess.Read.All` (no extra scope). Unknown kinds are
 * skipped — there is no known rules sub-route for them.
 */
type PolicyRuleKind = 'filtering' | 'tls' | 'prompt' | 'file' | 'cloudFirewall';

const RULE_ROUTE: Record<PolicyRuleKind, string> = {
  filtering: 'filteringPolicies',
  tls: 'tlsInspectionPolicies',
  prompt: 'promptPolicies',
  file: 'filePolicies',
  cloudFirewall: 'cloudFirewallPolicies',
};

/** Discriminate a policy link's rule kind from its (or its policy's) @odata.type. */
function policyRuleKind(link: FilteringPolicyLink): PolicyRuleKind | null {
  const linkType = link['@odata.type'];
  const policyType = (link.policy as { '@odata.type'?: string } | undefined)?.[
    '@odata.type'
  ];
  if (
    linkType === '#microsoft.graph.networkaccess.tlsInspectionPolicyLink' ||
    policyType === '#microsoft.graph.networkaccess.tlsInspectionPolicy'
  ) {
    return 'tls';
  }
  if (
    linkType === '#microsoft.graph.networkaccess.promptPolicyLink' ||
    policyType === '#microsoft.graph.networkaccess.promptPolicy'
  ) {
    return 'prompt';
  }
  if (
    linkType === '#microsoft.graph.networkaccess.filePolicyLink' ||
    policyType === '#microsoft.graph.networkaccess.filePolicy'
  ) {
    return 'file';
  }
  if (
    linkType === '#microsoft.graph.networkaccess.cloudFirewallPolicyLink' ||
    policyType === '#microsoft.graph.networkaccess.cloudFirewallPolicy'
  ) {
    return 'cloudFirewall';
  }
  if (
    linkType === '#microsoft.graph.networkaccess.filteringPolicyLink' ||
    policyType === '#microsoft.graph.networkaccess.webContentFilteringPolicy'
  ) {
    return 'filtering';
  }
  return null;
}

async function attachRulesByBatch(
  profiles: SecurityProfile[],
  signal?: AbortSignal,
): Promise<void> {
  // Map policy id → { writable policy object, rule kind }. A policy object may
  // hold any rule shape, so we attach as `unknown[]` (the row mapper narrows it
  // per @odata.type at display time).
  interface RuleTarget {
    policy: { policyRules?: unknown[] };
    kind: PolicyRuleKind;
  }
  const targets = new Map<string, RuleTarget>();

  for (const profile of profiles) {
    for (const link of profile.policies ?? []) {
      const kind = policyRuleKind(link);
      const policy = link.policy as ({ id?: string } & RuleTarget['policy']) | undefined;
      if (kind && policy?.id) {
        targets.set(policy.id, { policy, kind });
      }
    }
  }

  const uniqueIds = [...targets.keys()];
  if (uniqueIds.length === 0) return;

  const requests = uniqueIds.map((id) => ({
    id,
    method: 'GET' as const,
    url: `/networkAccess/${RULE_ROUTE[targets.get(id)!.kind]}/${id}/policyRules`,
  }));

  const responses = await graphBatch<ODataCollection<unknown>>(
    requests,
    { scopes: SCOPES, feature: FEATURE },
  );

  for (const [id, target] of targets) {
    const resp = responses.get(id);
    target.policy.policyRules =
      resp && resp.status >= 200 && resp.status < 300
        ? (resp.body.value ?? [])
        : [];
  }

  // AbortSignal is accepted for API symmetry; $batch does not stream per-item.
  void signal;
}

// ---------------------------------------------------------------------------
// What-If live user resolver (User.ReadBasic.All + GroupMember.Read.All)
// ---------------------------------------------------------------------------

const ENDPOINT_USER_BY_UPN = '/users?$filter=userPrincipalName eq {upn}';
const ENDPOINT_MEMBER_OF   = '/users/{id}/transitiveMemberOf';
const WHATIF_USER_FEATURE  = 'whatif-user-resolve';

export interface ResolvedUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  /** Transitive group ids (directory roles excluded) for CA group-membership evaluation. */
  groupIds: string[];
  /** Same groups with display names for the trace drawer. */
  groups: { id: string; displayName?: string }[];
}

/** Conservative UPN shape check run before the value reaches an OData filter. */
const UPN_RE = /^[A-Za-z0-9!#$%&*+\-/=?^_`{|}~.]{1,64}@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$/;

/**
 * A user-lookup failure the operator can act on. Named so the UI error mapper
 * can pass the message through instead of replacing it — these strings are
 * written here and carry no tenant data (see src/ui/friendlyError.ts).
 */
export class UserLookupError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'UserLookupError';
  }
}

/**
 * Resolve a UPN to a user object + transitive group ids for What-If group-targeted
 * CA evaluation. Uses the optional `User.ReadBasic.All` + `GroupMember.Read.All`
 * scopes — throws on missing consent or unknown UPN (caller surfaces inline error).
 */
export async function resolveUserByUpn(
  upn: string,
  signal?: AbortSignal,
): Promise<ResolvedUser> {
  const candidate = upn.trim();
  if (!UPN_RE.test(candidate)) {
    throw new UserLookupError(
      'Enter a valid user principal name, for example user@contoso.com.',
    );
  }

  // $filter is more reliable than /users/{upn} path lookup — avoids UPN encoding
  // issues. Single quotes are doubled per OData, then URLSearchParams encodes the
  // whole value, so the typed UPN can never break out of the string literal.
  const query = new URLSearchParams({
    $filter: `userPrincipalName eq '${candidate.replace(/'/g, "''")}'`,
    $select: 'id,displayName,userPrincipalName',
  });
  const filterResult = await graphGet<ODataCollection<{ id: string; displayName?: string; userPrincipalName?: string }>>(
    `/users?${query.toString()}`,
    { scopes: [USER_READ_SCOPE], feature: WHATIF_USER_FEATURE, endpoint: ENDPOINT_USER_BY_UPN, signal },
  );
  const user = filterResult.value?.[0];
  if (!user?.id) {
    throw new UserLookupError('No user found with that user principal name.');
  }

  const membership = await graphGet<ODataCollection<{ id: string; displayName?: string; '@odata.type'?: string }>>(
    `/users/${encodeURIComponent(user.id)}/transitiveMemberOf?$select=id,displayName`,
    { scopes: [GROUP_MEMBER_SCOPE], feature: WHATIF_USER_FEATURE, endpoint: ENDPOINT_MEMBER_OF, paginate: true, signal },
  );

  // Strip directory roles — keep only groups.
  const groups = (membership.value ?? [])
    .filter((o) => !(o['@odata.type'] ?? '').toLowerCase().includes('role'))
    .map((o) => ({ id: o.id, displayName: o.displayName }));
  const groupIds = groups.map((g) => g.id);

  return {
    id: user.id,
    displayName: user.displayName ?? upn,
    userPrincipalName: user.userPrincipalName ?? upn,
    groupIds,
    groups,
  };
}

