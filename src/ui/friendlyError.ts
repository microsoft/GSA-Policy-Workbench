/**
 * User-facing error text — findings 14 and 15 of the 2026-08-05 security review.
 *
 * Raw MSAL and Graph exception messages were rendered straight into the UI.
 * Those strings can embed the authority, correlation and tenant identifiers,
 * request URLs, and the value the operator typed — all of which then travel
 * into screenshots and support tickets.
 *
 * This maps the failure classes the app can actually produce onto plain
 * language, and falls back to a caller-supplied sentence for anything else.
 * Nothing derived from the error is interpolated into the result.
 *
 * Errors are matched by `name` and `errorCode` rather than by `instanceof`, so
 * this Tier-1 helper does not have to import the Graph adapter or MSAL.
 */

interface ErrorShape {
  name: string;
  message: string;
  /** Present on MSAL `AuthError` and its subclasses. */
  errorCode?: string;
  /** Present on the adapter's `GraphError`. */
  status?: number;
}

function shapeOf(err: unknown): ErrorShape {
  if (!(err instanceof Error)) return { name: '', message: '' };
  const withExtras = err as Error & { errorCode?: unknown; status?: unknown };
  return {
    name: err.name,
    message: err.message,
    errorCode: typeof withExtras.errorCode === 'string' ? withExtras.errorCode : undefined,
    status: typeof withExtras.status === 'number' ? withExtras.status : undefined,
  };
}

/** Graph HTTP status -> what the operator can do about it. */
function graphStatusText(status: number | undefined): string | null {
  if (status === undefined) return null;
  if (status === 401) return 'Your session has expired. Sign in again.';
  if (status === 403) {
    return 'Your account does not have permission to read this from Microsoft Graph.';
  }
  if (status === 404) return 'Microsoft Graph has no record matching that request.';
  if (status === 429) {
    return 'Microsoft Graph is throttling requests. Wait a moment and try again.';
  }
  if (status >= 500) {
    return 'Microsoft Graph is temporarily unavailable. Try again shortly.';
  }
  return null;
}

/** MSAL `errorCode` -> what the operator can do about it. */
const MSAL_TEXT: Record<string, string> = {
  user_cancelled: 'Sign-in was cancelled.',
  popup_window_error:
    'The sign-in window could not be opened. Allow pop-ups for this site and try again.',
  empty_window_error:
    'The sign-in window could not be opened. Allow pop-ups for this site and try again.',
  interaction_in_progress:
    'A sign-in is already in progress. Finish or close it, then try again.',
  consent_required: 'Additional consent is required. Sign in again to continue.',
  interaction_required: 'Additional consent is required. Sign in again to continue.',
  login_required: 'You are not signed in. Sign in and try again.',
  access_denied: 'Access was denied by your organisation’s policy.',
  invalid_grant: 'Your session is no longer valid. Sign in again.',
  no_network_connectivity:
    'Could not reach the sign-in service. Check your network connection.',
};

/**
 * Error classes whose messages this app writes itself. They are already plain
 * language and carry no tenant data, so they pass through unchanged.
 */
const OWN_ERROR_NAMES = new Set([
  'FixtureParseError',
  'InvalidTenantError',
  'UserLookupError',
]);

/**
 * Convert any thrown value into text safe to show the operator.
 *
 * @param fallback Sentence used when the failure is not one this app knows how
 *                 to explain. Should describe the action that failed.
 */
export function friendlyError(err: unknown, fallback: string): string {
  const { name, message, errorCode, status } = shapeOf(err);

  if (OWN_ERROR_NAMES.has(name)) return message;

  if (name === 'MissingScopeError') {
    return 'This needs a permission that has not been granted yet. Use “Enable full detail” to consent, then try again.';
  }

  if (name === 'UntrustedGraphUrlError') {
    return 'The request was blocked because it was addressed to an unexpected host.';
  }

  if (name === 'GraphError') {
    return graphStatusText(status) ?? fallback;
  }

  if (errorCode && MSAL_TEXT[errorCode]) return MSAL_TEXT[errorCode];

  return fallback;
}
