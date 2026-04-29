# Retyc Outlook Add-in — agent notes

## Mission

Outlook port of the Thunderbird extension at `../retyc-thunderbird-plugin`. Same end-to-end-encrypted file transfer
flow, identical UX strings where they apply, English-only code & comments.

**Do not invent new features.** When in doubt, mirror the Thunderbird behaviour and copy. The Outlook UX **does**
legitimately diverge in one place — see "Architectural choices" below.

## Versioning rule (read this before bumping anything)

Always pull the **latest non-deprecated** versions of every dependency. Never pin to old versions "for safety" — the
user pushes back on stale stacks.

The hard constraint is the **runtime baseline: Node 18.18.0+** (the user dev machine runs Node 18). When a package's
latest major requires Node 20+, drop one major and use the latest 1.x of that line. Do not pin to a Node 18 version of a
package that has compatible newer releases.

Procedure when adding/bumping a dep:

```bash
npm show <pkg> version              # what is latest?
npm show <pkg>@latest engines       # does latest support Node 18?
# if not:
npm show <pkg> versions             # find the highest version with Node 18 support
npm show <pkg>@<version> engines    # confirm
```

Currently downgraded for Node 18 compat (every other dep is on its absolute latest):

| Package                              | Pinned version | Why                                            |
|--------------------------------------|----------------|------------------------------------------------|
| `eslint`, `@eslint/js`               | `^9.39.4`      | v10 requires Node 20.19+ (`util.styleText`)    |
| `copy-webpack-plugin`                | `^13.0.1`      | v14 uses `Array.prototype.toSorted` (Node 20+) |
| `office-addin-manifest` (CI/release) | `1.13.6`       | v2 requires Node 20+ (ESM `strip-bom`)         |

When the user's Node baseline moves to ≥20.19, bump these to latest in the same commit and remove this table entry.

## Tech stack

- TypeScript 6 (strict, `noEmit`) bundled with webpack 5
- Vanilla DOM — no UI framework. Match Thunderbird's structure (no React, even though the original skeleton had it)
- `@retyc/sdk` (^1.0.1) for OIDC device flow, encryption, transfer upload
- Office.js, Mailbox API 1.7+ — only the compose-side `body.{getAsync,setAsync}` and `to.{getAsync,setAsync}` are used.
  **No event-based runtime, no Smart Alerts, no `displayDialogAsync`.**
- **Office Add-in XML manifest** — Classic Outlook desktop only accepts the legacy `<OfficeApp xsi:type="MailApp">`
  schema. The unified Microsoft 365 JSON manifest is *not* an option here.

## Repository layout

```
src/shared/   constants.ts, messages.ts (UserInfo only), settings.ts, sdk-factory.ts,
              token-store.ts, storage.ts, upload.ts (recipient parser + performRetycTransfer)
src/taskpane/ taskpane.{ts,html,css}  # ONLY runtime — auth + settings + dropzone + send-via-Retyc
src/types/    css.d.ts
assets/       icon-{16,32,48,64,80,128}.png  # PNG only — XML manifest rejects SVG
manifest.xml                          # canonical manifest, targets https://outlook.retyc.com (prod-ready)
manifest.dev.xml                      # generated for VM sideload, targets DEV_HOST (gitignored)
scripts/dev-setup.js                  # regenerates mkcert leaf + manifest.dev.xml from DEV_HOST
DEV_VM.md                             # dev procedure with a Windows VM (mkcert, sideload, networking)
Dockerfile + .docker/                 # static-asset hosting (multi-stage Node 22 → nginx 1.30 alpine-slim)
                                      #   .docker/nginx.conf
                                      #   .docker/40-substitute-base-url.sh   ← runtime BASE_URL substitution
.github/workflows/                    # _ci.yml + _docker.yml (reusable) called by main.yml + release.yml
```

## Architectural choices (and why)

- **Files are dropped INTO the task pane, never via Outlook's "Attach file" button.**.
  Reason: `getAttachmentContentAsync` is unreliable on Classic Outlook desktop — it returns empty
  content for attachments unless the draft has been saved to Exchange first, and forcing a `saveAsync` would push the
  unencrypted bytes to Exchange in cleartext, breaking E2EE. By bypassing Outlook's attachment system entirely (HTML5
  `File` from drag-drop or `<input type=file>`, then `file.arrayBuffer()`), the bytes go from disk → WebView2 → SDK
  encryption → Retyc, never touching Exchange.
- **Recipients are user-controlled in the task pane**, not polled from Outlook. We pre-fill the textarea once on load
  from `item.{to,cc,bcc}.getAsync` (one-shot, gated by a `_recipientsTouched` flag), and the user types/edits freely. On
  submit we mirror them into Outlook's `To` field via `item.to.setAsync(emails)` — Cc/Bcc are left untouched. A "Use
  Outlook recipients" button forces a re-sync.
- **There is no `OnMessageSend` handler.** Earlier iterations had one, but Office.js can't read attachment content from
  that event runtime on Classic desktop. The dropzone-first flow makes interception unnecessary.
- **No background script, no commands runtime, no dialog runtime.** Just the single task pane. Webpack has one entry 
  (`taskpane.ts`); the manifest declares one runtime (the task pane).
- **Persistence** travels through `Office.context.roamingSettings` (tokens, URLs, expiry). Wrapped by
  `src/shared/storage.ts` which falls back to `localStorage` (with `retyc:` prefix) when running outside Outlook so the
  task pane can be tested directly in a browser without crashing on `Office.context.roamingSettings === undefined`.
- **`OfficeRoamingTokenStore`** implements `TokenStore` from the SDK. Tokens are stored as a JSON-encoded string 
  (objects survive `set/get` but the SDK type validation prefers stringified JSON, and roaming-settings sync is more
  reliable on flat strings).
- **Body modification** uses `body.setAsync(existingHtml + linkSnippet)` — Office.js compose body has no `appendAsync`,
  so we read + concat + set. The OnMessageSend-only constraints (where only `appendOnSendAsync` works) don't apply here
  because we never modify the body from that runtime.
- **Passphrase OR recipients** — at least one is required. With recipients, they get decryption rights (their public key
  is fetched). With a passphrase ≥ 8 chars and no recipients, anyone with the link + passphrase can decrypt. The send
  button validates this combo in real time.
- **`manifest.xml` is prod-ready by default.** All URLs in the source point at `https://outlook.retyc.com`.
  `scripts/dev-setup.js` is the **single source of truth** for dev URL substitution — webpack does NOT rewrite URLs
  itself, it only picks which file to copy:
  - `npm run dev:setup` (override `DEV_HOST=...` as needed) → rewrites `outlook.retyc.com` → `${DEV_HOST}:3000` and
    writes the result to `manifest.dev.xml` (gitignored). Use `DEV_HOST=localhost` for direct localhost testing.
  - Webpack dev build (`npm run build:dev` / `npm run start`) → `pickManifestSource()` copies `manifest.dev.xml` to
    `dist/manifest.xml`. **Dev builds fail loudly if `manifest.dev.xml` is missing** (run `npm run dev:setup` first) —
    so we can't accidentally ship a `outlook.retyc.com`-pointing manifest to a local dev server.
  - Webpack prod build (`npm run build`, used inside the Docker image) → always copies `manifest.xml` so the entrypoint
    can substitute `BASE_URL`.
  - Docker → `.docker/40-substitute-base-url.sh` rewrites `outlook.retyc.com` → `${BASE_URL}` at container start
    (default kept at `https://outlook.retyc.com`, override via `-e BASE_URL=...`).

## Linter rules to remember

- `@typescript-eslint/no-floating-promises` — `Office.onReady` returns a Promise even with a callback. Prefix with
  `void`. Wrap async DOM listeners with `() => { void asyncFn() }`.
- `@typescript-eslint/no-unsafe-enum-comparison` would flag `Office.MailboxEnums.*` comparisons against strings. We
  currently don't compare any enum values — kept as a note in case future code reads attachment metadata or coercion
  types.
- ESLint flat config lives in `eslint.config.mjs` (the `.mjs` extension is mandatory because `package.json` is CommonJS
  and Node would refuse `import` syntax otherwise).

## Manifest gotchas (XML schema)

- Top-level `<Version>` must be **4-part** (e.g. `1.0.0.0`) AND `>= 1.0.0`. The validator rejects `0.x.y.z`.
- `<Permissions>ReadWriteItem</Permissions>` is sufficient — we only modify the active item (`to`, `body`). Don't bump
  to `ReadWriteMailbox` unless we genuinely need to read other emails.
- `<IconUrl>` / `<HighResolutionIconUrl>` must point to **PNG/JPG/GIF**, not SVG. The repo only carries PNG; if you need
  to regenerate from a vector source, use `inkscape --export-type=png --export-width=<size>` (input SVG is not
  committed).
- The compose button is in `<OfficeTab id="TabDefault">` (group label "Retyc"). On Classic desktop the user sees a "
  Retyc" group inside the Home tab and can reorder it via Customize the Ribbon. On new Outlook for Windows, custom tab
  placement is silently ignored — the button ends up in the "Apps" pivot regardless.
- `<SupportsPinning>true</SupportsPinning>` is set in the V1_1 `<Action xsi:type="ShowTaskpane">`. The task pane stays
  open across drafts when the user pins it.
- VersionOverrides V1_0 **and** V1_1 each declare their full Hosts/Resources blocks (graceful degradation — older
  clients read V1_0 without pinning, modern clients read V1_1 with pinning).
- Sideloading in Classic Outlook desktop **requires XML**. The "Add from File" picker filters on `.xml`. JSON unified
  manifests are silently rejected.
- Re-validate after every manifest edit:
  `npx -p office-addin-manifest@1.13.6 office-addin-manifest validate manifest.xml`.

## Common commands

```bash
npm run build:dev          # webpack dev build
npm run build              # production build
npm run watch              # incremental builds
npm run typecheck          # tsc --noEmit
npm run lint               # eslint src
npm run start              # webpack-dev-server (HTTPS, port 3000) — for sideloading
npm run dev:setup          # regenerate .certs/dev.{pem,-key.pem} + manifest.dev.xml (override DEV_HOST=...)
```

## Things to leave alone

- **The dropzone-first model.** Don't re-add `getAttachmentContentAsync` flows to read files attached via Outlook's 
  "Attach file" button — that path is unreliable on Classic Outlook desktop and requires `saveAsync`, which leaks the
  cleartext attachment to Exchange.
- **Don't re-introduce an `OnMessageSend` handler.** It can't read attachment content on Classic desktop, can't open the
  task pane (Microsoft restricts `commandId` to `executeFunction` actions, not `ShowTaskpane`), and adds two more
  runtimes' worth of plumbing. The dropzone replaces all of it.
- **Don't re-introduce a `displayDialogAsync` flow.** The task pane is the sole upload UI.
- **Don't poll Outlook's recipients periodically.** We pre-fill once + provide a manual "Use Outlook recipients" button.
  The task pane field is the source of truth on submit. (Earlier iterations polled every 2 s — removed because the UX
  of "Outlook says X but the pane shows Y" was confusing.)
- **Don't switch the SDK to ESM (`dist/index.js`).** The webpack alias targets `dist/index.cjs` to avoid bundling issues
  with Node-only sub-paths.
- **Don't add a ComposeView/ReadView/RetycLink scanner.** The original skeleton had stubs for those features; out of
  scope for the Thunderbird-parity goal.
- **Manifest UUID** is currently the placeholder `00000000-0000-0000-0000-000000000001`. **Replace with a real GUID
  before publishing to AppSource.**
- **Don't bake `BASE_URL` or `LANDING_URL` at Docker build time.** The image is environment-agnostic by design —
  both `manifest.xml` and the nginx config are kept as templates in `/etc/retyc/*.template` (outside the web root) and
  re-rendered at container start by `.docker/40-substitute-base-url.sh`. A single image serves preprod / prod / staging
  via `docker run -e BASE_URL=... -e LANDING_URL=...`. If you re-add a build-time substitution you lose that.
  - `BASE_URL` rewrites manifest URLs (taskpane + icons) so Outlook fetches from the right host.
  - `LANDING_URL` is the redirect target for the bare-domain `/` (only hit by users typing the asset host in a
    browser — Outlook itself loads `/taskpane.html` directly).

## Reference: Thunderbird sibling

`../retyc-thunderbird-plugin/` is the source of truth for UX strings, the SDK call shape, and the overall flow
philosophy. When implementing or fixing:

1. Find the equivalent in TB.
2. Adapt the API calls (Office.js `item.to.getAsync` instead of `browser.compose.getComposeDetails`, etc.).
3. Keep user-visible strings byte-identical when they describe the same UX moment. The Outlook task pane has its own
   strings for the dropzone (no equivalent in TB) — those are new and live here only.
