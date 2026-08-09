# Getting started

The primary way to use the Workbench is **signing in to a live tenant** — that's the whole point of the tool: one screen over your real GSA policy instead of nine admin-center blades. Loading an exported **policy file** is a secondary, sign-in-free path for a quick look with sample data or a previously-exported tenant.

## Prerequisites

- **Node.js 20+** and npm.
- For **live sign-in (the main path):** an Entra app registration (set up in step 2 below) and a GSA-licensed tenant in the Global ("WW") service deployment. The app hard-blocks US Gov / 21Vianet tenants.
- For **exporting your own policy file:** PowerShell 7+ and the Microsoft Graph PowerShell module (see `[testharness/README.md](../testharness/README.md)`).

## 1. Clone and install

```bash
git clone https://github.com/<org>/gsa-policy-workbench.git
cd gsa-policy-workbench
npm ci
```

## 2. App registration (optional — only if you want live sign-in now)

> Skip this step if you just want to try the Workbench first with the bundled sample policy file — jump to [step 3](#3-run-the-dev-server), choose **Load policy file** on the sign-in screen, and come back here when you're ready to point it at a real tenant. Since live sign-in is the primary way this tool is meant to be used, set this up before running the dev server so `.env.local` is in place on first launch.

The Workbench is a pure client-side SPA — it never sees your Graph token on a server, and it never requests write scopes. To sign in to a live tenant, you need an Entra app registration that grants the app permission to *read* GSA policy on your behalf.

### Required Entra role

The account you use to create the app registration and grant consent must be at least **Application Administrator** or **Cloud Application Administrator** in the target tenant.

### Delegated Graph scopes

All scopes below are **delegated** (the app only ever sees what the signed-in admin can see) and all are **admin-consent** scopes.

| Scope                    | Required?    | Unlocks                                                                                                                                           |
| ------------------------ | ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| `NetworkAccess.Read.All` | **Required** | The Security Profile / Filtering Policy / rule tree — the core of the app.                                                                        |
| `Policy.Read.All`        | Recommended  | Full Conditional Access targeting detail (users, groups, device filter) in the where-used drawer. Without it, CA policies show as name + ID only. |
| `Directory.Read.All`     | Optional     | Resolves user/group GUIDs to display names in CA targeting.                                                                                       |
| `Application.Read.All`   | Optional     | Private Access application segments and App Proxy detail in live mode.                                                                            |
| `User.ReadBasic.All`     | Optional     | What-If UPN lookup — resolves a typed UPN to a user object ID + display name.                                                                     |
| `GroupMember.Read.All`   | Optional     | What-If group-membership resolution — turns a group-targeted CA policy from "may apply" into a definitive yes/no for a specific user.             |

If you skip an optional scope, the corresponding feature degrades gracefully with an inline note in the UI — it does not break sign-in or the rest of the app.

### Option A — Automated (recommended)

`CreateAppRegistration.ps1` creates the app registration, requests the required scopes, grants admin consent, and writes `.env.local` in one step.

**Prerequisites:**

- PowerShell 7+ (`pwsh`). Verify with `pwsh --version`.
- The `Microsoft.Graph.Applications` and `Microsoft.Graph.Identity.SignIns` PowerShell modules:

  ```powershell
  Install-Module Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns -Scope CurrentUser
  ```

**Run it from the repo root:**

```powershell
pwsh -NoProfile -File ./CreateAppRegistration.ps1 -TenantId <your-tenant-id>
```

This opens a browser sign-in scoped to the tenant you specify, then:

1. Creates a **single-tenant SPA** app registration named `GSA Policy Workbench (local dev)` with redirect URI `http://localhost:5173`.
2. Requests the `NetworkAccess.Read.All` and `Policy.Read.All` delegated scopes.
3. Grants admin consent for those scopes (requires the Application/Cloud Application Administrator role above).
4. Writes `VITE_AAD_CLIENT_ID` and `VITE_AAD_TENANT` to `.env.local`.

Useful parameters:

| Parameter           | Default                            | Purpose                                                                                                                                                                                                       |
| ------------------- | ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `-TenantId`         | *(mandatory)*                      | The tenant to register the app in. The connection is always scoped to this tenant, so the app is guaranteed to land in the right place.                                                                       |
| `-Name`             | `GSA Policy Workbench (local dev)` | Display name for the app registration.                                                                                                                                                                        |
| `-RedirectUri`      | `http://localhost:5173`            | SPA redirect URI. Change this if you run the dev server on a different port.                                                                                                                                  |
| `-EnvPath`          | `.env.local`                       | Where to write the client ID / tenant ID. Use a per-tenant file (e.g. `.env.contoso.local`) if you test against multiple tenants — see [Testing against multiple tenants](#testing-against-multiple-tenants). |
| `-Force`            | off                                | Overwrite the env file if it already exists.                                                                                                                                                                  |
| `-SkipAdminConsent` | off                                | Create the app without granting consent — useful if you aren't a tenant admin; hand the client ID to one to consent manually.                                                                                 |

The optional scopes in the table above (`Directory.Read.All`, `Application.Read.All`, `User.ReadBasic.All`, `GroupMember.Read.All`) are **not** requested by the script today. Add them by hand in the Entra admin center (Option B, step 2) if you want the features they unlock.

### Option B — Manual

In the [Entra admin center](https://entra.microsoft.com), under **Identity → Applications → App registrations → New registration**:

| Field                   | Value                                                              |
| ----------------------- | ------------------------------------------------------------------ |
| Name                    | `GSA Policy Workbench (local dev)` (or anything you like)          |
| Supported account types | **Accounts in this organizational directory only** (single tenant) |
| Redirect URI — platform | **Single-page application (SPA)**                                  |
| Redirect URI — value    | `http://localhost:5173`                                            |

After creation, on the app registration page:

1. **Authentication** → confirm `http://localhost:5173` is listed under the SPA platform. Leave "Allow public client flows" off — this app never needs it.
2. **API permissions** → Add a permission → Microsoft Graph → **Delegated** → add the scopes you want from the table above.
3. **API permissions** → **Grant admin consent for `<tenant>**`. All scopes here are admin-consent; skipping this step means sign-in works but consent fails.
4. **Overview** → copy the **Application (client) ID** and the **Directory (tenant) ID**.

Then write `.env.local` by hand:

```bash
cp .env.example .env.local   # PowerShell: Copy-Item .env.example .env.local
```

```env
VITE_AAD_CLIENT_ID=<paste your Application (client) ID>
VITE_AAD_TENANT=<paste your Directory (tenant) ID>
```

`VITE_AAD_TENANT` defaults to `organizations` (multi-tenant, tenant picked at sign-in) when unset. For single-tenant dev, paste the actual tenant GUID so MSAL pins the authority correctly.

> `VITE_AAD_CLIENT_ID` is a **public client identifier**, not a secret — it's fine to have it in a client-side bundle. Never add a client *secret* to this app; a SPA app registration has none, and adding one would be a credential-leak risk in the shipped bundle.

### Testing against multiple tenants

Each tenant needs its own app registration (single-tenant by design). Vite's `--mode` flag lets you keep one env file per tenant:

```powershell
pwsh -NoProfile -File ./CreateAppRegistration.ps1 -TenantId <tenant-id> -EnvPath .env.<name>.local
```

```bash
npm run dev -- --mode <name>   # loads .env.<name>.local
npm run dev                    # loads .env.local (default)
```

All `.env.*.local` files are git-ignored — they never leave your machine.

## 3. Run the dev server

```bash
npm run dev
```

Open the printed URL (default [http://localhost:5173](http://localhost:5173)). You land on the sign-in screen, which offers both paths below.

## 4a. Path A — Live sign-in (the main path)

1. Make sure step 2 is done — `.env.local` has your client ID and tenant ID.
2. On the sign-in screen, pick the tenant and **Sign in with Microsoft**, then grant admin consent if prompted. Commercial (worldwide) tenants only — GovCloud / 21Vianet are blocked by design.

## 4b. Path B — Offline file mode (no sign-in, quick look)

1. On the sign-in screen, choose **Load policy file** instead.
2. Select an exported policy JSON file. To try it immediately with sanitized data, use the bundled `[Config-Sample/sample-tenant.json](../Config-Sample/sample-tenant.json)`.
3. The full Internet Access and Private Access views render with no tenant, sign-in, or network access — the file is served through the same Tier 3 code path as live mode, via a swappable transport seam.

To inspect **your own** tenant offline instead of live, export one first (see [Exporting your own tenant policy file](#5-exporting-your-own-tenant-policy-file-optional)).

## 5. Exporting your own tenant policy file (optional)

To build a policy file from a live tenant for offline (file-mode) use, run the test harness:

```powershell
cd testharness
.\Export-GsaFixture.ps1 -TenantId "<tenant-guid>" -Verbose
# → ./exports/gsa-fixture_<timestamp>.json  — your exported policy file (load it via "Load policy file")
```

It captures the Internet Access profile tree plus the Private Access domain (Private Access / Quick Access apps, App Proxy apps, Conditional Access, and authentication-strength policies) as a single JSON document. Full options and required scopes are in `[testharness/README.md](../testharness/README.md)`.

> ⚠️ Exports contain **real** tenant data and are git-ignored — never commit a raw export. Share only sanitized files such as `[Config-Sample/sample-tenant.json](../Config-Sample/sample-tenant.json)`.

## Useful scripts

| Command             | What it does                                               |
| ------------------- | ---------------------------------------------------------- |
| `npm run dev`       | Start the Vite dev server (hot reload).                    |
| `npm run build`     | Type-check and produce a production build.                 |
| `npm run preview`   | Serve the production build locally.                        |
| `npm run typecheck` | Strict TypeScript check over `src/` and `tests/`, no emit. |
| `npm run test`      | Run the test suite (`tests/`) with Vitest.                 |
| `npm run lint`      | ESLint over the project.                                   |
| `npm run audit`     | Fail on high/critical advisories in shipped dependencies.  |
