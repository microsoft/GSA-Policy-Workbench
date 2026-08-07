/**
 * Tenant identifier validation — finding 13 of the 2026-08-05 security review.
 *
 * The tenant is runtime-supplied (the field on the sign-in card, or
 * `VITE_AAD_TENANT`) and is interpolated directly into the MSAL authority URL.
 * A value containing `/` or `..` rewrites the authority path, so it is checked
 * against the three shapes Entra actually accepts before it is used.
 *
 * Kept separate from `msalConfig.ts` so the check can be tested without
 * constructing a `PublicClientApplication` or touching `window`.
 */

/** Directory (tenant) ID. */
const TENANT_GUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Verified domain such as `contoso.onmicrosoft.com` — DNS labels only. */
const TENANT_DOMAIN_RE =
  /^[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$/;

/**
 * Meta-tenants this app accepts. `consumers` is deliberately excluded: every
 * scope this app requests needs admin consent in a work or school tenant.
 */
const META_TENANTS: readonly string[] = ['common', 'organizations'];

/** A DNS name cannot exceed this, and neither can a GUID. */
const MAX_TENANT_LENGTH = 253;

export class InvalidTenantError extends Error {
  constructor() {
    // The rejected value is deliberately not interpolated — it is untrusted
    // input and would end up in screenshots and support tickets.
    super(
      'Enter a tenant ID, a verified domain such as contoso.onmicrosoft.com, ' +
        "or 'organizations'.",
    );
    this.name = 'InvalidTenantError';
  }
}

/** True when `tenant` is a directory ID, a verified domain, or a meta-tenant. */
export function isValidTenant(tenant: string): boolean {
  const value = tenant.trim();
  if (value === '' || value.length > MAX_TENANT_LENGTH) return false;
  return (
    META_TENANTS.includes(value.toLowerCase()) ||
    TENANT_GUID_RE.test(value) ||
    TENANT_DOMAIN_RE.test(value)
  );
}

/** Returns the trimmed tenant, or throws `InvalidTenantError`. */
export function assertValidTenant(tenant: string): string {
  const value = tenant.trim();
  if (!isValidTenant(value)) throw new InvalidTenantError();
  return value;
}
