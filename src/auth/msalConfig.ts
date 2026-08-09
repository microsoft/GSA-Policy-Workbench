import {
  PublicClientApplication,
  LogLevel,
  type Configuration,
} from '@azure/msal-browser';
import { assertValidTenant } from './tenant';

const clientId = import.meta.env.VITE_AAD_CLIENT_ID;
const defaultTenant = import.meta.env.VITE_AAD_TENANT || 'organizations';

/**
 * Build the authority URL for a given tenant. The tenant may be a directory ID,
 * a verified domain, or one of the meta-tenants ("organizations", "common").
 *
 * Throws `InvalidTenantError` on anything else, so a value carrying `/` or `..`
 * can never rewrite the authority path. A bad `VITE_AAD_TENANT` therefore fails
 * at startup rather than at the first sign-in.
 */
export function authorityFor(tenant: string): string {
  return `https://login.microsoftonline.com/${assertValidTenant(tenant)}`;
}

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: authorityFor(defaultTenant),
    // Allow signing into any tenant the user picks at runtime.
    knownAuthorities: [],
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Warning,
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        if (level === LogLevel.Error) console.error(message);
      },
    },
  },
};

export const pca = new PublicClientApplication(msalConfig);

export const DEFAULT_TENANT = defaultTenant;
