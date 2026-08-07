/**
 * Graph delegated scopes used by the Workbench.
 *
 * `NetworkAccess.Read.All` is required. All others are optional and degrade
 * gracefully when absent (no interactive popup mid-load).
 *
 * `Policy.Read.All`      — full CA targeting conditions; degrades to stubs.
 * `Directory.Read.All`   — resolves user/group GUIDs to display names.
 * `Application.Read.All` — Private Access segments + App Proxy (live mode).
 * `User.ReadBasic.All`   — What-If UPN lookup: resolve a typed UPN to user id + displayName.
 * `GroupMember.Read.All` — What-If group membership: resolve group-targeted CA policies
 *                          from "may apply" to a definitive yes/no for a specific user.
 *
 * All scopes require admin consent. See docs/architecture.md §10 and getting-started.md §1.
 */
export const REQUIRED_SCOPES = ['NetworkAccess.Read.All'] as const;

export const OPTIONAL_SCOPES = [
  'Policy.Read.All',
  'Directory.Read.All',
  'Application.Read.All',
  'User.ReadBasic.All',
  'GroupMember.Read.All',
] as const;

export const ALL_SCOPES = [...REQUIRED_SCOPES, ...OPTIONAL_SCOPES] as const;

/** Scope string used to detect whether the CA-detail capability is granted. */
export const CA_DETAIL_SCOPE = 'Policy.Read.All';

/** Scope used to resolve user/group object ids to display names (optional). */
export const DIRECTORY_SCOPE = 'Directory.Read.All';

/**
 * Scope used to read Application Proxy apps and Private Access application
 * objects (segments, onPremisesPublishing). Optional — the Private Access
 * view degrades to showing SPs without segment detail when absent.
 * See spec.md §6.5.
 */
export const APPLICATION_SCOPE = 'Application.Read.All';

/** What-If UPN resolver: look up a user by UPN to get their object id + displayName. */
export const USER_READ_SCOPE = 'User.ReadBasic.All';

/** What-If group membership: resolve transitive groups so group-targeted CA policies evaluate to yes/no. */
export const GROUP_MEMBER_SCOPE = 'GroupMember.Read.All';
