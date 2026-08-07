# GSA Policy Workbench

> A single-pane policy management workbench for **Microsoft Global Secure Access (GSA) Internet Access** and **Private Access** — unified, searchable, read-only visibility across the policy blades that today are scattered across 9–10 screens in the Entra admin center.

**Status:** V1 — full read-only visibility plane (Internet Access + Private Access, live tenant or an exported policy file).

---

## Why this exists

Administering Microsoft GSA policy today means jumping across nearly a dozen separate blades in the Entra admin center, with no single screen that shows every rule, in order, searchable and filterable, alongside the Conditional Access policies that target it.

The Workbench is a **read-only inspector**: point it at a tenant (or an exported policy file) and it renders the full policy tree — Security Profiles, Filtering Policies, rules, Conditional Access targeting, and Private Access apps — in one searchable, filterable table. It does not write anything to Graph in V1: it is a visibility tool, not a policy editor or migration tool.

It is **not** a policy generator, migration tool, or multi-cloud tool. If you're looking for those, see [Related tools](#related-tools) below.

## Features

- 🔐 **MSAL sign-in** — an Entra-styled sign-in card with a tenant picker and a Global-service (non-GovCloud) guard. US Gov and 21Vianet tenants are hard-blocked, not degraded.
- 📂 **Offline file-based data source** — load an exported policy file and inspect it with no sign-in and no network access, useful for demos, support cases, or air-gapped review.
- 🧩 **Single Graph adapter** — every Graph call flows through one adapter (`src/adapters/graph/`), with every call passing through an audit interceptor so a full call log is always available for troubleshooting.
- 📥 **Adaptive profile-tree loader** — issues the deep `$expand` in one call and, if the tenant's Graph beta surface won't return rules inline, transparently falls back to a batched per-policy fetch.
- 🏠 **Home overview** — a planner-styled landing screen with a **What-If** effective-policy resolver (rule-match only, not a live-traffic simulator).
- 📋 **Internet Access policy table** — every Security Profile → linked Filtering Policy → rule, in collapsible per-profile sections with search, filters, and linked Conditional Access targeting.
- 🔒 **Private Access policy view** — Private Access / Quick Access apps and Application Proxy apps in a grouped table that mirrors the Internet Access look, correlated to covering Conditional Access policies (including App Proxy pre-authentication posture).
- 🔗 **Conditional Access targeting** — CA → Security Profile / app correlation with user/group targeting detail where `Policy.Read.All` is consented, and graceful degradation to name + ID where it is not.

> ⚠️ **Rule ordering note.** The GSA Graph beta models priority on Security Profiles and on policy-links within a profile, but **filtering rules inside a policy have no order field**. The table sorts by profile and policy-link priority; rule rows render first-match-as-returned and are explicitly **not** reorderable. This matches the API reality and is intentional — the Workbench will not fake a drag-and-drop reorder UI that Graph can't back.

## Related tools

- **Migrating from a third-party SSE product?** Use [Migrate2GSA](https://aka.ms/migrate2gsa) — this project does not import third-party configuration.
- **Building a greenfield policy set from personas?** Use the [EIA Greenfield Wizard](https://microsoft.github.io/Migrate2GSA/tools/EIA-Greenfield-Wizard/) — this project does not generate policy.

---

## Getting started

The primary way to use the Workbench is **signing in to a live tenant** — that's the whole point of the tool: one screen over your real GSA policy instead of nine admin-center blades. Loading an exported **policy file** is a secondary, sign-in-free path for a quick look with sample data or a previously-exported tenant.

### Prerequisites

- **Node.js 20+** and npm.
- For **live sign-in (the main path):** an Entra app registration (set up in step 2 below) and a GSA-licensed tenant in the Global ("WW") service deployment. The app hard-blocks US Gov / 21Vianet tenants.
- For **exporting your own policy file:** PowerShell 7+ and the Microsoft Graph PowerShell module (see `[testharness/README.md](testharness/README.md)`).

### 1. Clone and install

```bash
git clone https://github.com/<org>/gsa-policy-workbench.git
cd gsa-policy-workbench
npm ci
```

### 2. App registration (optional — only if you want live sign-in now)

> Skip this step if you just want to try the Workbench first with the bundled sample policy file — jump to [step 3](#3-run-the-dev-server), choose **Load policy file** on the sign-in screen, and come back here when you're ready to point it at a real tenant. Since live sign-in is the primary way this tool is meant to be used, set this up before running the dev server so `.env.local` is in place on first launch.

The Workbench is a pure client-side SPA — it never sees your Graph token on a server, and it never requests write scopes. To sign in to a live tenant, you need an Entra app registration that grants the app permission to *read* GSA policy on your behalf.

#### Required Entra role

The account you use to create the app registration and grant consent must be at least **Application Administrator** or **Cloud Application Administrator** in the target tenant.

#### Delegated Graph scopes

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

#### Option A — Automated (recommended)

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

#### Option B — Manual

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

```
VITE_AAD_CLIENT_ID=<paste your Application (client) ID>
VITE_AAD_TENANT=<paste your Directory (tenant) ID>
```

`VITE_AAD_TENANT` defaults to `organizations` (multi-tenant, tenant picked at sign-in) when unset. For single-tenant dev, paste the actual tenant GUID so MSAL pins the authority correctly.

> `VITE_AAD_CLIENT_ID` is a **public client identifier**, not a secret — it's fine to have it in a client-side bundle. Never add a client *secret* to this app; a SPA app registration has none, and adding one would be a credential-leak risk in the shipped bundle.

#### Testing against multiple tenants

Each tenant needs its own app registration (single-tenant by design). Vite's `--mode` flag lets you keep one env file per tenant:

```powershell
pwsh -NoProfile -File ./CreateAppRegistration.ps1 -TenantId <tenant-id> -EnvPath .env.<name>.local
```

```bash
npm run dev -- --mode <name>   # loads .env.<name>.local
npm run dev                    # loads .env.local (default)
```

All `.env.*.local` files are git-ignored — they never leave your machine.

### 3. Run the dev server

```bash
npm run dev
```

Open the printed URL (default [http://localhost:5173](http://localhost:5173)). You land on the sign-in screen, which offers both paths below.

### 4a. Path A — Live sign-in (the main path)

1. Make sure step 2 is done — `.env.local` has your client ID and tenant ID.
2. On the sign-in screen, pick the tenant and **Sign in with Microsoft**, then grant admin consent if prompted. Commercial (worldwide) tenants only — GovCloud / 21Vianet are blocked by design.

### 4b. Path B — Offline file mode (no sign-in, quick look)

1. On the sign-in screen, choose **Load policy file** instead.
2. Select an exported policy JSON file. To try it immediately with sanitized data, use the bundled `[Config-Sample/sample-tenant.json](Config-Sample/sample-tenant.json)`.
3. The full Internet Access and Private Access views render with no tenant, sign-in, or network access — the file is served through the same Tier 3 code path as live mode, via a swappable transport seam.

To inspect **your own** tenant offline instead of live, export one first (see [Exporting your own tenant policy file](#5-exporting-your-own-tenant-policy-file-optional)).

### 5. Exporting your own tenant policy file (optional)

To build a policy file from a live tenant for offline (file-mode) use, run the test harness:

```powershell
cd testharness
.\Export-GsaFixture.ps1 -TenantId "<tenant-guid>" -Verbose
# → ./exports/gsa-fixture_<timestamp>.json  — your exported policy file (load it via "Load policy file")
```

It captures the Internet Access profile tree plus the Private Access domain (Private Access / Quick Access apps, App Proxy apps, Conditional Access, and authentication-strength policies) as a single JSON document. Full options and required scopes are in `[testharness/README.md](testharness/README.md)`.

> ⚠️ Exports contain **real** tenant data and are git-ignored — never commit a raw export. Share only sanitized files such as `[Config-Sample/sample-tenant.json](Config-Sample/sample-tenant.json)`.

### Useful scripts

| Command             | What it does                                               |
| ------------------- | ---------------------------------------------------------- |
| `npm run dev`       | Start the Vite dev server (hot reload).                    |
| `npm run build`     | Type-check and produce a production build.                 |
| `npm run preview`   | Serve the production build locally.                        |
| `npm run typecheck` | Strict TypeScript check over `src/` and `tests/`, no emit. |
| `npm run test`      | Run the test suite (`tests/`) with Vitest.                 |
| `npm run lint`      | ESLint over the project.                                   |
| `npm run audit`     | Fail on high/critical advisories in shipped dependencies.  |

### Where to go next

- **First-run flow and troubleshooting:** [docs/getting-started.md](docs/getting-started.md)
- **How the UI is laid out:** [docs/UX-Design-Patterns.md](docs/UX-Design-Patterns.md)
- **How the layers fit together:** [docs/architecture.md](docs/architecture.md)
- **What the user sees and does:** [docs/spec.md](docs/spec.md)

---

## Deploying

> **Recommendation: host on Azure Static Web Apps, or on your own web server.**  
> Use GitHub Pages only for a sign-in-free, file-mode-only build.

This is a static SPA with no backend, so it will run from any static host. The hosts are **not** equivalent from a security standpoint, because the app holds Microsoft Graph access tokens in the browser.

### Azure Static Web Apps, or customer-hosted — recommended

`[staticwebapp.config.json](staticwebapp.config.json)` ships the Content Security Policy and browser security headers. Azure Static Web Apps applies it automatically. On your own web server, reproduce the same headers in the server config.

Two properties make this the right default:

- **The full header set applies**, including `frame-ancestors` (clickjacking protection), `X-Content-Type-Options`, `Referrer-Policy`, and `Permissions-Policy`. None of these can be set from inside a static bundle.
- **The app gets its own origin.** MSAL keeps the token cache in `sessionStorage`, which is scoped per origin — so a dedicated origin means no other application can read this app's Graph tokens.

Set the redirect URI in the app registration to the deployed origin, and confirm the effective response headers against the live URL before release.

### GitHub Pages — offline builds only

GitHub Pages **cannot set response headers at all**. A Content Security Policy can only be delivered from a `<meta http-equiv>` tag in [index.html](index.html), and that route has hard limits — `frame-ancestors` is [not supported in `<meta>](https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Headers/Content-Security-Policy/frame-ancestors)`, so the page can be framed by any origin, and `X-Content-Type-Options` and `Permissions-Policy` cannot be delivered at all.

There is also an origin-sharing problem: a project site is served from `https://<org>.github.io/<repo>/`, a **single origin shared by every Pages site under that account**. Because `sessionStorage` is scoped per origin, any other site published there could in principle read this app's MSAL token cache.

Neither issue matters if the deployment never signs in, so GitHub Pages is a reasonable home for a **file-mode-only build** — the offline inspector, loading an exported policy file, holding no tokens. Do not publish a sign-in-capable build there.

Two things to get right when hosting under a repository subpath:

- Set Vite's `[base](https://vite.dev/config/shared-options.html#base)` to the subpath, or every asset 404s.
- The MSAL redirect URI must be the full subpath URL, registered as-is. The origin alone is not enough.

---

## Architecture

- **Frontend:** Vite + React + TypeScript (strict)
- **Component library:** Fluent UI v9 (chrome + cells)
- **Table engine:** TanStack Table (headless) + TanStack Virtual — **not** Fluent `DataGrid`
- **Auth:** MSAL.js (delegated scopes, tenant-pickable)
- **Data layer:** TanStack Query over a single Graph adapter
- **No backend**

The **Graph adapter** (`src/adapters/graph/`) is the key abstraction. Every Graph call goes through it, and every call passes through an audit interceptor that records the endpoint template (never populated URLs — no tenant, user, or policy identifiers). Do not call Graph from components; see [docs/architecture.md](docs/architecture.md) for the full three-tier design.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

[MIT](LICENSE).
