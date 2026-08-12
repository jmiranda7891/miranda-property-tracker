# CLAUDE.md — Miranda Family Property Tracker

Portable project memory. Chat history does NOT travel between sessions — this file + the repo are the only bridge. Open this first on any machine or in any new session.

## What this is
A private estate-planning web app listing the Miranda family's ~40 real estate properties across Mexico and the Chicago area — unified list (cards or a sortable Excel-like table), Google Maps links, one-time AI-researched area context + value-estimate commentary per property, an EN/ES toggle, a Google Sheet export, and an in-app admin panel for managing who can access it. Built for Jorge (jorge@orderexpress.com) as a personal/family tool — not a client product, not related to the sibling `clpr-social-media` (CL Social Media App) repo, though it deliberately reuses that repo's proven Apps Script patterns.

## Current status
- **v1** (version 1/2 on the deployment): initial build — unified property list, Maps links, bilingual research/value text, EN/ES toggle, export, admin-managed access.
- **v1.1** (version 3, current): added a second view — a sortable, Excel-like table (`STATE.view`, `renderListView()`, `sortProperties()` in `JavaScript.html`) alongside the original card grid, plus a "subtle/professional" card refresh (colored left-border accent by status, a compact MX/US country chip, a one-click "View on Maps" link on each card). Purely client-side — no server or data-schema changes. 75 tests passing (up from 67).

## Hosting / accounts
- Lives under the **OE (Order Express)** Google Workspace account, domain `orderexpress.com`.
- Web app deployed domain-restricted (`access: "DOMAIN"` in `appsscript.json`), but that is only the OUTER gate — the real access control is the app's own `Users` allow-list (see below). An `orderexpress.com` account that isn't on the list gets a clean "no access yet" screen, not the app.
- Apps Script project: script id in `.clasp.json`. Single deployment, no separate dev/exec tiers (see "Deploy" below).
- **Live URL**: `https://script.google.com/a/macros/orderexpress.com/s/AKfycbxx-MBnRft8Zp7xYQkCmsR7Yg0VxeojqLrcj5XjyNjiM95ZUT3Ayj9ALhhLfGqUdGA/exec`

## Architecture
- Apps Script web app: `doGet` → HtmlService. Client↔server ONLY via `google.script.run` (async).
- Files: `Code.js` (server), `Index.html` (shell), `Styles.html`, `JavaScript.html` (client app). Assembled with `include()`. Exactly these 5 files (+`appsscript.json`) are ever pushed — `.claspignore` is a whitelist, not a blocklist.
- Data: **one spreadsheet total** — "Miranda Family - Property Registry" — created lazily by `ensureRegistry_()` the first time anyone opens the app, ID stored in Script Properties (`REGISTRY_ID`), never hardcoded. This is a SINGLE-TENANT app (one family, one property list), unlike the sibling CL app's per-client-spreadsheet pattern — there is no "Clients" tab or provisioning flow here.
  - `Properties` tab: unified MX+USA schema (`countryCode`, `paisRaw`, `tipo`, `referencia`, `propietario`, `direccion`, `ciudad`, `estado`, `cp`, `observacionesRaw`, `status`, `precioEstimadoUSD`, `precioEstimadoIsPlaceholder`, `escrituras`/`propEscriturado`/`propuestaTraspaso` [MX-only], `pin` [USA-only], `aiResearchEN`/`ES`, `aiValueEstimateEN`/`ES`, `researchDate`, `archived`/`archivedReason`, `createdBy`/`createdAt`, `updatedBy`/`updatedAt`).
  - `Users` tab: `email`, `role` (`member` | `admin`), `addedBy`, `addedAt`.
  - `AuditLog` tab: `ts`, `user`, `action`, `entityType`, `entityId`, `detail` (JSON — full row snapshot for `delete_property`, since hard delete needs a recovery path).
- No secrets in Script Properties for v1 — no Gemini key, no Maps API key. There is NO live AI integration: the `aiResearch*`/`aiValueEstimate*` fields are static text, hand-authored once (see "Data & research" below), not generated at runtime.
- Google Maps links are computed **client-side only**, from address+city+state+country via `encodeURIComponent` into a plain `google.com/maps/search` URL — no API key, no geocoding, ever.

## Access control
- `roleFor_(email)`: if `Users` is empty, the first person to open the app becomes `admin` (bootstrap). Otherwise, an email not on the list gets role `'none'` — **not** a "guest" role, this is a private app, not org-open. `requireUser_()` gates all reads; `requireAdmin_()` gates admin-only writes.
- `member`: view everything, add a new property, edit any property's fields.
- `admin`: everything a member can, plus archive/unarchive/delete a property, and manage the `Users` list (add/remove/change role). Same "can't demote/remove the last admin" and "can't remove yourself" guards as the sibling app.
- Delete is a true destructive row-delete (admin-only). The client makes the admin type the property's `referencia` to confirm; the server checks it again (`deleteProperty(id, confirmReferencia)`); the full row is snapshotted into `AuditLog.detail` first, since that snapshot is the only recovery path.

## Data & research (read this before touching `SEED_PROPERTIES_`)
- `SEED_PROPERTIES_` in `Code.js` is a hand-authored array of all 41 properties (34 Mexico + 7 USA), cleaned up from the family's original Excel listing (`Mexico` and `USA` tabs only — the `Mexico Propuesta 2` and `Sheet3` tabs in that file are historical/redundant and were never migrated).
- Cleanup decisions baked into the seed data: country spelling normalized (`Mexico`/`México` → `countryCode: 'MX'`), `status` derived from the original `Observaciones` text (kept verbatim in `observacionesRaw` too), the one placeholder price (`Oficina OE Maravatio`, literally `"1"` in the source) flagged via `precioEstimadoIsPlaceholder: true` rather than treated as a real number, "Oficina Aurora"'s city corrected from "Chicago" to "Aurora, IL" to match its actual address, and the sold Chicago house ("casa Willow") pre-archived (`archived: true, status: 'sold'`).
- `aiResearchEN/ES` and `aiValueEstimateEN/ES` were written **once**, by an AI assistant doing real web research per unique city/area (not per individual property — properties sharing a city reuse that area's research, personalized per property type/status) and are stored as plain static text. **There is no "regenerate" button and no runtime AI call anywhere in this app** — that was a deliberate v1 scope decision (Jorge: "I rather you do the research now... we won't need to update constantly"). If the research ever needs refreshing, it's a manual re-write of the relevant `SEED_PROPERTIES_` entries (or a direct edit in the underlying Sheet — the UI shows these fields read-only), not a feature to build.
- **The registry seeds itself automatically.** `ensureRegistry_()` calls `seedIntoRegistry_()` the moment the spreadsheet is first created — i.e., the very first time anyone opens the app. There is no manual "run this once from the Apps Script editor" step, because a headless `clasp`-deployed cloud session has no browser to click Run in, and `clasp run` requires the project to be set up as an API-executable (extra GCP wiring not worth doing for a one-time seed). `seedProperties_()` still exists as an admin-callable fallback, guarded by the same `SEED_DONE` Script Property, in case the registry ever needs a manual re-seed.
- **`SEED_DEFAULTS_`** supplies blank defaults for fields most rows don't need (`propietario`, `cp`, `escrituras`, `pin`, etc.) — each `SEED_PROPERTIES_` entry only needs to specify fields that differ from blank/false.

## Deploy workflow
- **Single deployment, no dev/exec two-tier ceremony** (unlike the sibling CL app) — this is a small family tool with a handful of users, not a team product where protecting a shared URL from an in-progress push matters. One deployment id, reused via `clasp deploy -i "$WEBAPP_DEPLOYMENT_ID" -V <version> -d "description"`.
- `./clasp-push.sh` pushes the current working tree to `@HEAD` and byte-verifies the push landed (`clasp_verify_push` — see the sibling CL app's CLAUDE.md for the 2026-08-06 incident this defends against: a push can print "Pushed N files" and exit 0 while the deployed script stays stale).
- To actually ship a change to the live `/exec` URL after pushing to `@HEAD`:
  ```
  clasp -A .clasprc.json create-version "description of the change"   # prints "Created version N"
  clasp -A .clasprc.json deploy -i "$WEBAPP_DEPLOYMENT_ID" -V N -d "description of the change"
  ```
  Note clasp v3's `create-version` takes the description **positionally**, not via `-d` (that flag is silently wrong on `create-version`, though correct on `deploy`) — same trap documented in the sibling repo.
- `clasp-env.sh` (sourced by `clasp-push.sh`) restores `.clasprc.json` from `CLASP_AUTH_OE_B64` if it's missing, finds clasp without requiring a global install, and refuses to run outside this repo's directory. Always pass `-A .clasprc.json` explicitly on every clasp command in a cloud session — bare `clasp` looks for a global `~/.clasprc.json` that doesn't exist there.
- `CLASP_AUTH_OE_B64` was already present in this environment when this project was built (proven working for several other OE-account repos: `oe-monitoring-app`, `oepm-trackerbot`, `oe-jcarlos-app`, `loft-dashboard-appscript`) — no separate auth setup was needed.
- Rollback: `clasp -A .clasprc.json deploy -i "$WEBAPP_DEPLOYMENT_ID" -V <older version> -d "rollback to vX"` — versions are immutable, so old ones are always still there.

## Apps Script gotchas (same class of bug as the sibling repo — read before editing served files)
- **Comment-stripping**: HtmlService's sanitizer can corrupt a `/* */` sequence even where it isn't a real comment (e.g. inside a string literal like `accept='image/*'`). The safest rule, followed throughout this repo: **`Index.html`, `Styles.html`, and `JavaScript.html` contain the two-character sequence `/*` or `*/` NOWHERE at all** — not even in legitimate CSS section comments. `test/static.js` enforces this with a whole-file scan (stricter than the sibling repo's suite, which only checks the served `<script>` block and exempts `Styles.html` — that exemption may be correct per that repo's own test comment, but this repo doesn't rely on it either way). `Code.js` is NOT served through HtmlService and can use normal `/* */` comments freely.
- **Load-order**: never read data at top level in `JavaScript.html`. The client calls `getBootstrap()` on load and computes everything inside that success handler.
- **Single spreadsheet, execution-scoped cache**: `CACHE_` in `Code.js` avoids re-opening the registry spreadsheet multiple times within one execution (Apps Script rebuilds globals fresh on every invocation, so this cache can never go stale ACROSS requests — only kills repeated `openById` calls WITHIN one). `bustReg_()` clears it after any write.
- **`ensureRegistry_()` recursion trap**: `seedIntoRegistry_()` is called from inside `ensureRegistry_()`, before `CACHE_.ss` is set — it must never call `ss_()`/`audit_()` (which would call `ensureRegistry_()` again and create a second spreadsheet). It writes directly to the `ss` object it's passed instead. Keep this in mind if `seedIntoRegistry_()` or `ensureRegistry_()` is ever refactored.

## Testing
- `npm test` runs `test/all.js`: `static` (parses + the comment-stripping scan + `.claspignore` whitelist + oauth scope checks + seed-data shape checks — 29 checks), `server` (the REAL `Code.js` evaluated in a bare `vm` context with fake Apps Script services from `test/lib/gas-stubs.js` — RBAC, property CRUD, admin user-management guards, audit logging, export, the auto-seed-on-bootstrap behavior — 28 checks), `client` (the REAL `Styles.html`/`JavaScript.html` in Chromium via Playwright, with `google.script.run` mocked — access-denied screen, filtering, the EN/ES toggle switching chrome+research text while never translating raw property data, admin-only UI hidden from a member — 10 checks). 67 checks total, ~3s.
- `test/lib/harness.js`, `test/lib/gas-stubs.js` ported/adapted from the sibling CL app's `test/lib/harness.js` pattern (`test/lib/gas-stubs.js` is new here — the sibling's server suite didn't need a fake spreadsheet since it only tested pure functions; this app's server logic is mostly sheet reads/writes, so `gas-stubs.js` adds an in-memory fake `SpreadsheetApp`/`Sheet`/`Range`).
- What the suite CANNOT cover: real Sheets/Drive quota, real Google sign-in, or the actual OE Workspace domain restriction. A manual pass as an authorized OE user against the live `/exec` URL is the only way to verify those.

## `appsscript.json` scopes
Minimal for what v1 actually does: `spreadsheets`, `drive.file` (narrower than a sibling-style full `drive` scope — this app only ever creates its OWN spreadsheets via `SpreadsheetApp.create`, for the registry and for exports, never opens arbitrary Drive folders), `userinfo.email` (`Session.getActiveUser().getEmail()`). No `script.external_request` (no `UrlFetchApp`), no `script.send_mail` (no email), no `script.scriptapp` (no triggers) — v1 makes no external HTTP calls and sets no time-driven triggers.

## Pending / next (not built in v1, by design)
- **No live AI integration** — see "Data & research" above. If this ever needs to change (e.g. genuinely wanting a "refresh this property's research" button calling Gemini), it's new scope: a `GEMINI_KEY` Script Property, a new oauth scope (`script.external_request`), and a new server function — not a small tweak to the existing static fields.
- **No map with pins** — only plain "open in Google Maps" links (no API key, per Jorge's explicit "skip to plain links for now"). A pinned map view would need a billed Google Maps JavaScript API key.
- **AI research/value text is read-only in the UI** — editable only by directly editing the underlying Sheet, not through the app.
- Everything else asked for in the original brief (unified list, Maps links, bilingual AI research/value text, EN/ES toggle, Sheet export, admin-managed access, add/archive/delete permissions split by role) shipped in v1.

## Setup on a new machine / session (clone-based)
1. Clone: `git clone https://github.com/jmiranda7891/miranda-property-tracker.git`
2. `npm install` (installs Playwright for the `client` test suite only — nothing here is deployed to Apps Script; `.claspignore` is the whitelist that enforces that).
3. In a cloud session, `CLASP_AUTH_OE_B64` should already be present in the environment; `clasp-env.sh` restores `.clasprc.json` from it automatically on the first `./clasp-push.sh` run.
4. `npm test` before any push — it's fast (~3s) and it's what would catch a broken deploy before Apps Script does.
5. `./clasp-push.sh` to push to `@HEAD`; see "Deploy workflow" above to actually move the live `/exec` URL to a new version.
6. Nothing secret to provision — the registry spreadsheet is created automatically (and seeds itself) the first time an authorized user opens the app. Do not manually create a "Miranda Family - Property Registry" spreadsheet — that would race with the app's own `ensureRegistry_()` and potentially produce a duplicate.
