/**
 * Fixture parsing for file-based data source.
 *
 * A fixture file is a Graph-wire response body: literally what
 * `GET /networkAccess/filteringProfiles?$expand=…` returns. Providing
 * `policyRules` inline on each policy exercises the loader's Branch A.
 *
 * This parser is intentionally shallow — it validates the top-level OData
 * collection shape only. The generic mapper (mapper.ts) still picks the
 * declared fields downstream, so deep validation here would be redundant.
 */

import type { ODataCollection } from '../../model/types';
import type { SecurityProfile } from '../../model/definitions/SecurityProfile.definition';
import type {
  ConditionalAccessPolicy,
  DirectoryObjectRef,
} from '../../model/definitions/ConditionalAccessPolicy.definition';
import type {
  PrivateAccessApp,
  AppProxyApp,
  AuthenticationStrengthPolicy,
} from '../../model/definitions/PrivateAccessApp.definition';
import type { ForwardingProfile } from '../../model/definitions/ForwardingProfile.definition';
import type { TenantPolicy } from '../graph/loader';

/** Thrown when a fixture file is not valid or has an unexpected shape. */
export class FixtureParseError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'FixtureParseError';
  }
}

/** Guard rails for untrusted input. A fixture is a hand-supplied file. */
const MAX_FIXTURE_BYTES = 64 * 1024 * 1024;
const MAX_COLLECTION_ITEMS = 100_000;
const FORBIDDEN_KEYS = new Set(['__proto__', 'prototype', 'constructor']);

/**
 * Reject prototype-polluting keys anywhere in the parsed document. `JSON.parse`
 * creates `__proto__` as an own property rather than invoking the setter, but any
 * later spread or merge would promote it, so refuse the file outright.
 */
function assertNoForbiddenKeys(node: unknown, depth = 0): void {
  if (depth > 64) throw new FixtureParseError('File is nested too deeply.');
  if (Array.isArray(node)) {
    for (const item of node) assertNoForbiddenKeys(item, depth + 1);
    return;
  }
  if (node === null || typeof node !== 'object') return;
  for (const key of Object.keys(node)) {
    if (FORBIDDEN_KEYS.has(key)) {
      throw new FixtureParseError(
        `File contains a disallowed property name ("${key}").`,
      );
    }
    assertNoForbiddenKeys((node as Record<string, unknown>)[key], depth + 1);
  }
}

/**
 * Validate an optional top-level sibling collection: must be an array of plain
 * objects and within the item cap. Returns `undefined` when absent.
 */
function optionalCollection<T>(
  obj: Record<string, unknown>,
  key: string,
): T[] | undefined {
  const raw = obj[key];
  if (raw === undefined) return undefined;
  if (!Array.isArray(raw)) {
    throw new FixtureParseError(`Expected '${key}' to be an array.`);
  }
  if (raw.length > MAX_COLLECTION_ITEMS) {
    throw new FixtureParseError(`'${key}' exceeds ${MAX_COLLECTION_ITEMS} items.`);
  }
  for (const item of raw) {
    if (item === null || typeof item !== 'object' || Array.isArray(item)) {
      throw new FixtureParseError(`'${key}' must contain only objects.`);
    }
  }
  return raw as T[];
}

/**
 * The `GET /networkAccess/filteringProfiles?$expand=…` response body.
 *
 * The test-harness exporter shapes fixtures with the full Conditional Access
 * policy detail as a top-level sibling of `value`. Live Graph does not return
 * this sibling — the SPA only sees CA stubs (id + displayName) on each
 * profile — so the loader degrades gracefully when it is absent.
 */
export interface FilteringProfilesBody extends ODataCollection<SecurityProfile> {
  conditionalAccessPolicies?: ConditionalAccessPolicy[];
  /** Resolved user/group display names for CA targeting (exporter only). */
  directoryObjects?: DirectoryObjectRef[];
  /** Referenced authentication-strength policies (exporter only — spec §6.3). */
  authenticationStrengthPolicies?: AuthenticationStrengthPolicy[];
  /** Tagged Private Access / Quick Access service principals (exporter only). */
  privateAccessApps?: PrivateAccessApp[];
  /** Application Proxy applications with onPremisesPublishing (exporter only). */
  appProxyApps?: AppProxyApp[];
  /** GSA traffic-forwarding profiles (internet / m365 / private) with state. */
  forwardingProfiles?: ForwardingProfile[];
  /** Tenant-wide policy inventory (id + name + type) for orphaned detection. */
  tenantPolicies?: TenantPolicy[];
}

/**
 * A parsed fixture document. Today it holds a single response — the expanded
 * Security Profile tree. An envelope keyed by endpoint template is an open
 * design decision (open decision #1).
 */
export interface FixtureDoc {
  /** Body for `GET /networkAccess/filteringProfiles?$expand=…`. */
  filteringProfiles: FilteringProfilesBody;
}

/**
 * Parse and validate fixture file text. Throws `FixtureParseError` with a
 * human-readable message on any problem, so the file picker can surface it.
 */
export function parseFixture(text: string): FixtureDoc {
  if (text.length > MAX_FIXTURE_BYTES) {
    throw new FixtureParseError(
      `File is larger than ${Math.round(MAX_FIXTURE_BYTES / 1024 / 1024)} MB.`,
    );
  }

  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    throw new FixtureParseError('File is not valid JSON.');
  }

  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
    throw new FixtureParseError('Expected a JSON object at the top level.');
  }

  assertNoForbiddenKeys(raw);

  const obj = raw as Record<string, unknown>;
  const value = obj['value'];
  if (!Array.isArray(value)) {
    throw new FixtureParseError(
      "Expected an OData collection with a 'value' array " +
        '(a GET /networkAccess/filteringProfiles response).',
    );
  }
  if (value.length > MAX_COLLECTION_ITEMS) {
    throw new FixtureParseError(`'value' exceeds ${MAX_COLLECTION_ITEMS} items.`);
  }
  for (const item of value) {
    if (item === null || typeof item !== 'object' || Array.isArray(item)) {
      throw new FixtureParseError("'value' must contain only profile objects.");
    }
  }

  // The exporter attaches full CA detail as a top-level sibling of `value`.
  // All siblings are optional — a fixture without them is a valid V0 fixture.
  const conditionalAccessPolicies =
    optionalCollection<ConditionalAccessPolicy>(obj, 'conditionalAccessPolicies');
  const directoryObjects =
    optionalCollection<DirectoryObjectRef>(obj, 'directoryObjects');
  const authenticationStrengthPolicies =
    optionalCollection<AuthenticationStrengthPolicy>(obj, 'authenticationStrengthPolicies');
  const privateAccessApps =
    optionalCollection<PrivateAccessApp>(obj, 'privateAccessApps');
  const appProxyApps = optionalCollection<AppProxyApp>(obj, 'appProxyApps');
  const forwardingProfiles =
    optionalCollection<ForwardingProfile>(obj, 'forwardingProfiles');
  const tenantPolicies = optionalCollection<TenantPolicy>(obj, 'tenantPolicies');

  return {
    filteringProfiles: {
      value: value as SecurityProfile[],
      // `@odata.nextLink` is deliberately dropped: the file transport answers any
      // filteringProfiles path from this same document, so honouring a next link
      // would loop the paginator forever.
      conditionalAccessPolicies,
      directoryObjects,
      authenticationStrengthPolicies,
      privateAccessApps,
      appProxyApps,
      forwardingProfiles,
      tenantPolicies,
    },
  };
}
