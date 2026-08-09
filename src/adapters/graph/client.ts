import { recordGraphCall } from './interceptor';
import { getGraphFetch } from './transport';

const GRAPH_BASE =
  import.meta.env.VITE_GRAPH_BASE_URL ?? 'https://graph.microsoft.com/beta';

const API_VERSION: 'beta' | 'v1.0' = GRAPH_BASE.includes('/v1.0')
  ? 'v1.0'
  : 'beta';

/**
 * Hosts allowed to receive an access token. Anything else — including a
 * tampered `@odata.nextLink` or a mis-set `VITE_GRAPH_BASE_URL` — is rejected
 * before the Authorization header is attached.
 */
const ALLOWED_GRAPH_HOSTS: readonly string[] = ['graph.microsoft.com'];

/** Pagination safety limits — a hostile or looping nextLink must not hang the tab. */
const MAX_PAGES = 500;
const MAX_ITEMS = 250_000;

export class UntrustedGraphUrlError extends Error {
  constructor(readonly url: string) {
    // The rejected URL is deliberately not interpolated — it may carry tenant data.
    super('Refusing to send an access token to a non-Graph URL.');
    this.name = 'UntrustedGraphUrlError';
  }
}

/**
 * True when `raw` is a Graph URL this app may send a bearer token to: HTTPS,
 * an allowlisted host, the default port, no embedded credentials, and a
 * recognised API-version path prefix.
 */
export function isAllowedGraphUrl(raw: string): boolean {
  let u: URL;
  try {
    u = new URL(raw);
  } catch {
    return false;
  }
  return (
    u.protocol === 'https:' &&
    ALLOWED_GRAPH_HOSTS.includes(u.hostname) &&
    u.port === '' &&
    u.username === '' &&
    u.password === '' &&
    /^\/(beta|v1\.0)(\/|$)/.test(u.pathname)
  );
}

function assertAllowedGraphUrl(raw: string): string {
  if (!isAllowedGraphUrl(raw)) throw new UntrustedGraphUrlError(raw);
  return raw;
}

// Fail closed at startup rather than at the first request.
if (!isAllowedGraphUrl(`${GRAPH_BASE.replace(/\/$/, '')}/`)) {
  throw new Error(
    'VITE_GRAPH_BASE_URL must be an https://graph.microsoft.com/{beta|v1.0} URL.',
  );
}

/** Async function that returns an access token for the given scopes, or null. */
export type TokenProvider = (scopes: string[]) => Promise<string | null>;

let tokenProvider: TokenProvider | null = null;

/** Registered by the AuthProvider. The adapter is auth-agnostic otherwise. */
export function setTokenProvider(provider: TokenProvider): void {
  tokenProvider = provider;
}

/**
 * When true, auth is bypassed entirely — `authHeader()` returns a placeholder
 * and never calls the token provider. Set by the file-based data source so the
 * app can run with no tenant and no sign-in.
 */
let authBypassed = false;

export function setAuthBypass(bypass: boolean): void {
  authBypassed = bypass;
}

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly body?: unknown,
  ) {
    super(message);
    this.name = 'GraphError';
  }
}

export class MissingScopeError extends Error {
  constructor(readonly scope: string) {
    super(`Access token for required scope "${scope}" is unavailable.`);
    this.name = 'MissingScopeError';
  }
}

interface GraphRequestOptions {
  scopes: string[];
  /** Logical feature for the gap report, e.g. 'load-table'. */
  feature: string;
  /** Endpoint template for the gap report (no ids). Defaults to the path. */
  endpoint?: string;
  /** Follow @odata.nextLink and concatenate `value` arrays. */
  paginate?: boolean;
  signal?: AbortSignal;
  /**
   * Additional HTTP headers merged into the request.
   * Use for Graph advanced-query requirements such as `ConsistencyLevel: eventual`
   * (needed for `$filter` on `tags/any(...)`).
   */
  extraHeaders?: Record<string, string>;
  /**
   * Override the API version for this specific call. Defaults to the version
   * embedded in `GRAPH_BASE` (beta). Pass `'v1.0'` for endpoints that require
   * v1.0 (e.g. `/servicePrincipals` tag filter).
   */
  apiVersion?: 'beta' | 'v1.0';
}

async function authHeader(scopes: string[]): Promise<string> {
  if (authBypassed) return 'Bearer file-mode';
  if (!tokenProvider) throw new Error('No token provider registered.');
  const token = await tokenProvider(scopes);
  if (!token) throw new MissingScopeError(scopes[0] ?? 'unknown');
  return `Bearer ${token}`;
}

function toUrl(pathOrUrl: string): string {
  return pathOrUrl.startsWith('http')
    ? pathOrUrl
    : `${GRAPH_BASE}${pathOrUrl.startsWith('/') ? '' : '/'}${pathOrUrl}`;
}

/**
 * Like `toUrl` but respects an explicit `apiVersion` override. When the
 * requested version differs from `API_VERSION`, the version segment in
 * `GRAPH_BASE` is substituted. Used by `graphGet` to route calls to v1.0
 * endpoints (e.g. `/servicePrincipals`) while keeping the default base as beta.
 */
function resolveUrl(pathOrUrl: string, apiVersion?: 'beta' | 'v1.0'): string {
  if (pathOrUrl.startsWith('http')) return pathOrUrl;
  const base = apiVersion && apiVersion !== API_VERSION
    ? GRAPH_BASE.replace(/\/(beta|v1\.0)$/, `/${apiVersion}`)
    : GRAPH_BASE;
  return `${base}${pathOrUrl.startsWith('/') ? '' : '/'}${pathOrUrl}`;
}

const GUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** A path segment safe to keep verbatim: a literal route or OData/type token. */
const SAFE_SEGMENT_RE = /^[$]?[A-Za-z][A-Za-z0-9._]*$/;

/**
 * Reduce a populated request path to an audit-safe endpoint *template*.
 *
 * The interceptor stores templates only — never the populated URL (no tenant,
 * user, or policy identifiers). Callers normally
 * pass an explicit `endpoint` template, but this is the **fail-closed** default
 * so a caller that forgets one can never leak an id into the audit log or the
 * shareable call-log download.
 *
 * Allow-list rather than deny-list: a segment is kept only when it looks like a
 * literal route name. GUIDs, numeric ids, UPNs, encoded values, and any other
 * unrecognised shape collapse to `{id}`.
 *
 * Exported for the security regression suite (`tests/security/client.test.ts`).
 */
export function templatePath(pathOrUrl: string): string {
  return pathOrUrl
    .split('?')[0]
    .split('/')
    .map((seg) => {
      if (seg === '') return seg;
      if (GUID_RE.test(seg)) return '{id}';
      return SAFE_SEGMENT_RE.test(seg) ? seg : '{id}';
    })
    .join('/');
}

/**
 * Low-level GET. When `paginate` is set and the response is a collection,
 * follows `@odata.nextLink` until exhausted and returns the merged collection.
 */
export async function graphGet<T>(
  path: string,
  opts: GraphRequestOptions,
): Promise<T> {
  const authorization = await authHeader(opts.scopes);
  const endpoint = opts.endpoint ?? templatePath(path);
  const callApiVersion = opts.apiVersion ?? API_VERSION;

  let url: string | null = resolveUrl(path, opts.apiVersion);
  let merged: unknown[] | null = null;
  let firstBody: Record<string, unknown> | null = null;
  let pages = 0;
  const seen = new Set<string>();

  while (url) {
    assertAllowedGraphUrl(url);
    if (seen.has(url)) {
      throw new GraphError(`GET ${endpoint} returned a cyclic nextLink.`, 0);
    }
    seen.add(url);
    if (++pages > MAX_PAGES) {
      throw new GraphError(
        `GET ${endpoint} exceeded the ${MAX_PAGES}-page limit.`,
        0,
      );
    }

    const started = performance.now();
    const res = await getGraphFetch()(url, {
      headers: { authorization, accept: 'application/json', ...opts.extraHeaders },
      signal: opts.signal,
    });
    const durationMs = Math.round(performance.now() - started);

    recordGraphCall({
      ts: new Date().toISOString(),
      method: 'GET',
      endpoint,
      scopes: opts.scopes,
      status: res.status,
      durationMs,
      apiVersion: callApiVersion,
      feature: opts.feature,
    });

    if (!res.ok) {
      const body = await safeJson(res);
      throw new GraphError(
        `GET ${endpoint} failed: ${res.status}`,
        res.status,
        body,
      );
    }

    const body = (await res.json()) as Record<string, unknown>;
    firstBody ??= body;

    const value = body['value'];
    if (opts.paginate && Array.isArray(value)) {
      merged ??= [];
      // Appended one at a time: `push(...value)` passes every element as an
      // argument and overflows the stack on a large page, which would crash
      // before the MAX_ITEMS ceiling below could reject it.
      for (const item of value) merged.push(item);
      if (merged.length > MAX_ITEMS) {
        throw new GraphError(
          `GET ${endpoint} exceeded the ${MAX_ITEMS}-item limit.`,
          0,
        );
      }
      const next = body['@odata.nextLink'];
      url = typeof next === 'string' && next !== '' ? next : null;
    } else {
      return body as T;
    }
  }

  // Paginated collection: return a synthetic collection body.
  return { ...(firstBody as object), value: merged ?? [] } as T;
}

/**
 * Low-level POST with a JSON body (e.g. `directoryObjects/getByIds`). The
 * endpoint template is recorded by the interceptor; ids travel in the body,
 * never the URL, so the audit log stays templated.
 */
export async function graphPost<T>(
  path: string,
  body: unknown,
  opts: GraphRequestOptions,
): Promise<T> {
  const authorization = await authHeader(opts.scopes);
  const endpoint = opts.endpoint ?? templatePath(path);
  const started = performance.now();
  const res = await getGraphFetch()(assertAllowedGraphUrl(toUrl(path)), {
    method: 'POST',
    headers: {
      authorization,
      'content-type': 'application/json',
      accept: 'application/json',
    },
    body: JSON.stringify(body),
    signal: opts.signal,
  });
  const durationMs = Math.round(performance.now() - started);

  recordGraphCall({
    ts: new Date().toISOString(),
    method: 'POST',
    endpoint,
    scopes: opts.scopes,
    status: res.status,
    durationMs,
    apiVersion: API_VERSION,
    feature: opts.feature,
  });

  if (!res.ok) {
    const errBody = await safeJson(res);
    throw new GraphError(`POST ${endpoint} failed: ${res.status}`, res.status, errBody);
  }

  return (await res.json()) as T;
}

interface BatchRequest {
  id: string;
  method: 'GET';
  url: string;
  /** Per-request headers merged into the $batch body (e.g. ConsistencyLevel: eventual). */
  headers?: Record<string, string>;
}

interface BatchResponseItem<T> {
  id: string;
  status: number;
  body: T;
}

/**
 * Issue a JSON batch (max 20 requests per batch per Graph limits). Splits
 * larger request sets into multiple batches automatically.
 */
export async function graphBatch<T>(
  requests: BatchRequest[],
  opts: Pick<GraphRequestOptions, 'scopes' | 'feature'>,
): Promise<Map<string, BatchResponseItem<T>>> {
  const authorization = await authHeader(opts.scopes);
  const results = new Map<string, BatchResponseItem<T>>();

  for (let i = 0; i < requests.length; i += 20) {
    const chunk = requests.slice(i, i + 20);
    const started = performance.now();
    const res = await getGraphFetch()(assertAllowedGraphUrl(toUrl('/$batch')), {
      method: 'POST',
      headers: {
        authorization,
        'content-type': 'application/json',
        accept: 'application/json',
      },
      body: JSON.stringify({
        requests: chunk.map(r => ({
          id: r.id,
          method: r.method,
          url: r.url,
          ...(r.headers ? { headers: r.headers } : {}),
        })),
      }),
    });
    const durationMs = Math.round(performance.now() - started);

    recordGraphCall({
      ts: new Date().toISOString(),
      method: 'POST',
      endpoint: '/$batch',
      scopes: opts.scopes,
      status: res.status,
      durationMs,
      apiVersion: API_VERSION,
      feature: opts.feature,
    });

    if (!res.ok) {
      const body = await safeJson(res);
      throw new GraphError(`$batch failed: ${res.status}`, res.status, body);
    }

    const body = (await res.json()) as {
      responses: BatchResponseItem<T>[];
    };
    for (const item of body.responses ?? []) {
      results.set(item.id, item);
    }
  }

  return results;
}

async function safeJson(res: Response): Promise<unknown> {
  try {
    return await res.json();
  } catch {
    return undefined;
  }
}
