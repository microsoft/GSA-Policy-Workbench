# GSA Policy Workbench

> A single-pane policy management workbench for **Microsoft Global Secure Access (GSA) Internet Access** and **Private Access** — unified, searchable, read-only visibility across the policy blades that today are scattered across 9–10 screens in the Entra admin center.

---

## What is the GSA Policy Workbench tool?

Administering Microsoft GSA policy today means jumping across nearly a dozen separate blades in the Entra admin center, with no single screen that shows every rule, in order, searchable and filterable, alongside the Conditional Access policies that target it.

The Workbench is a **read-only inspector**: point it at a tenant (or an exported policy file) and it renders the full policy tree — Security Profiles, Filtering Policies, rules, Conditional Access targeting, and Private Access apps — in one searchable, filterable table. It does not write anything to Graph in V1: it is a visibility tool, not a policy editor or migration tool.

It is **not** a policy generator, migration tool, or multi-cloud tool. If you're looking for those, see [Related tools](#related-tools) below.

## Is this an official Microsoft product?

No. It is a community project maintained by Microsoft employees. The Workbench is provided as-is and is not supported through any Microsoft support program or service. Please do not contact Microsoft support with any issues or concerns.

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

```bash
git clone https://github.com/<org>/gsa-policy-workbench.git
cd gsa-policy-workbench
npm ci
npm run dev
```

Open the printed URL (default [http://localhost:5173](http://localhost:5173)) and choose **Load policy file** on the sign-in screen to try it immediately with the bundled `[Config-Sample/sample-tenant.json](Config-Sample/sample-tenant.json)` — no sign-in or network access required.

To point the Workbench at a live tenant, set up an Entra app registration first. Full prerequisites, the app registration walkthrough (automated script or manual), running the dev server, both sign-in paths, exporting your own tenant policy file, and the useful-scripts reference all live in **[docs/getting-started.md](docs/getting-started.md)**.

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

The **Graph adapter** (`src/adapters/graph/`) is the key abstraction. Every Graph call goes through it, and every call passes through an audit interceptor that records the endpoint template (never populated URLs — no tenant, user, or policy identifiers). Do not call Graph from components.

> **Note on automated tests.** This project's automated test suite hasn't been published to this repo yet. If you'd like us to share it, please open a GitHub Issue and vote for it.

## How to contribute

Please create a GitHub Issue to discuss the changes you are planning to make, then send us a PR.

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit <https://cla.opensource.microsoft.com>.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

[MIT](LICENSE).
