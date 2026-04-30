# Retyc Outlook Add-in — agent notes

## Mission

Outlook task pane that uploads files to Retyc with end-to-end encryption and replaces them by a secure download link
in the message body. English-only code & comments.

**Do not invent new features.** Stay within the scope described in "Architectural choices" below.

## Versioning rule (read this before bumping anything)

Always pull the **latest non-deprecated** versions of every dependency. Never pin to old versions "for safety" — the
user pushes back on stale stacks. Runtime baseline is **Node 24+** (declared in `package.json` engines and `.nvmrc`),
so there's no compat ceiling to worry about — `npm show <pkg> version && npm install <pkg>@latest`.

## Tech stack

- TypeScript 6 (strict, `noEmit` via `vue-tsc`) bundled with **Vite 8 + Rolldown**
- **Vue 3 + Nuxt UI v4** standalone (no Nuxt — we configure the `@nuxt/ui/vite` plugin directly with
  `router: false` / `colorMode: false`). Tailwind v4 with `@theme static` for the `cornflower-blue` / `jacaranda` palettes.
- `vue-i18n@11` in composition mode (`legacy: false`) — locale auto-detected from
  `Office.context.displayLanguage` then `navigator.language`, with EN / FR bundled.
- `@retyc/sdk` for OIDC device flow, encryption, transfer upload.
- Office.js, Mailbox API 1.7+ — only the compose-side `body.{getAsync,setAsync}`, `to.{getAsync,setAsync}` and the
  `RecipientsChanged` / `ItemChanged` events are used. **No event-based runtime, no Smart Alerts, no
  `displayDialogAsync`.**
- **Office Add-in XML manifest** — Classic Outlook desktop only accepts the legacy `<OfficeApp xsi:type="MailApp">`
  schema. The unified Microsoft 365 JSON manifest is *not* an option here.

## Repository layout

```
src/main.ts                           # bootstrap: Buffer polyfill, lucide icons, i18n, Office.onReady → mount
src/App.vue                           # auth gate + UTabs (Transfer / Account)
src/components/                       # LoginWall, TransferTab, AccountTab, DropZone (all .vue, lazy-loaded)
src/composables/                      # useAuth, useTransfer, useDropOverlay
src/shared/                           # constants.ts (API_URL, STORAGE_KEY_TOKENS), sdk-factory.ts,
                                      # storage.ts (roamingSettings + localStorage fallback), token-store.ts,
                                      # upload.ts (recipient parser + performRetycTransfer)
src/i18n.ts + src/locales/            # vue-i18n singleton + en/fr message bundles
src/assets/custom.css                 # @theme static block for the custom palette (no global resets)
src/shims/fs.ts                       # empty fs shim, aliased in vite.config.mts for browser bundles
src/env.d.ts                          # ImportMetaEnv typing for VITE_RETYC_API_URL
public/                               # static assets served as-is (icons, etc.)
taskpane.html                         # Vite entry HTML (kept at repo root — moving it caused issues)
manifest.xml                          # prod manifest, targets https://outlook.retyc.com
manifest.dev.xml                      # generated for VM sideload, targets DEV_HOST (gitignored)
                                      # other manifest.*.xml files (e.g. internal envs) are also gitignored
scripts/dev-setup.js                  # regenerates mkcert leaf + manifest.dev.xml from DEV_HOST
DEV_VM.md                             # dev procedure with a Windows VM (mkcert, sideload, networking)
Dockerfile + .docker/                 # static-asset hosting (multi-stage Node 24 → nginx 1.30 alpine-slim)
                                      #   .docker/nginx.conf
.github/workflows/                    # _ci.yml + _docker.yml (reusable) called by main.yml + release.yml
```

## Architectural choices (and why)

- **Files are dropped INTO the task pane, never via Outlook's "Attach file" button.**.
  Reason: `getAttachmentContentAsync` is unreliable on Classic Outlook desktop — it returns empty
  content for attachments unless the draft has been saved to Exchange first, and forcing a `saveAsync` would push the
  unencrypted bytes to Exchange in cleartext, breaking E2EE. By bypassing Outlook's attachment system entirely (HTML5
  `File` from drag-drop or `<input type=file>`, then `file.arrayBuffer()`), the bytes go from disk → WebView2 → SDK
  encryption → Retyc, never touching Exchange.
- **Outlook is the single source of truth for recipients.** The pane shows them as read-only `UBadge`s — there's no
  pane-side recipient editor. `useTransfer.ts` reads `item.{to,cc,bcc}.getAsync` once on mount, then keeps the list in
  sync via two Office.js handlers: `RecipientsChanged` on the active item (re-reads when the user edits To/Cc/Bcc) and
  `ItemChanged` on the mailbox (drops the previous item's `RecipientsChanged` handler before re-attaching to the new
  draft — failing to do so leaks handlers across drafts). The user's own address is filtered out via `isSelf()`. On
  submit, the merged deduplicated list is written back into Outlook's `To` via `item.to.setAsync(emails)` — Cc/Bcc are
  left untouched.
- **There is no `OnMessageSend` handler.** Earlier iterations had one, but Office.js can't read attachment content from
  that event runtime on Classic desktop. The dropzone-first flow makes interception unnecessary.
- **No background script, no commands runtime, no dialog runtime.** Just the single task pane. Vite has one entry
  (`taskpane.html` → `src/main.ts`); the manifest declares one runtime (the task pane).
- **Persistence** carries only the OIDC tokens, through `Office.context.roamingSettings` (mailbox-scoped, syncs across
  devices). API URL is build-time only (`VITE_RETYC_API_URL`, default `https://api.retyc.com`) — there is no
  user-adjustable settings UI. `src/shared/storage.ts` falls back to `localStorage` (with `retyc:` prefix) when running
  outside Outlook so the task pane can be tested directly in a browser without crashing on
  `Office.context.roamingSettings === undefined`.
- **`OfficeRoamingTokenStore`** implements `TokenStore` from the SDK. Tokens are stored as a JSON-encoded string 
  (objects survive `set/get` but the SDK type validation prefers stringified JSON, and roaming-settings sync is more
  reliable on flat strings).
- **Body modification** uses `body.setAsync(existingHtml + linkSnippet)` — Office.js compose body has no `appendAsync`,
  so we read + concat + set. The OnMessageSend-only constraints (where only `appendOnSendAsync` works) don't apply here
  because we never modify the body from that runtime.
- **Passphrase OR recipients** — at least one is required. With recipients, they get decryption rights (their public key
  is fetched). With a passphrase ≥ 8 chars and no recipients, anyone with the link + passphrase can decrypt. The send
  button validates this combo in real time.
- **Per-environment Docker images, no runtime substitution.** Each environment builds its own image: the manifest is
  baked in (`manifest.xml` for prod; for any other env, drop a gitignored `manifest.<env>.xml` and copy it over
  `manifest.xml` before `docker build`) and the API URL comes from the build arg `VITE_RETYC_API_URL` (defaults to
  `https://api.retyc.com` if unset, see `src/shared/constants.ts`). The Vite plugin only chooses the source file:
  - `npm run dev:setup` (override `DEV_HOST=...` as needed) → rewrites `outlook.retyc.com` → `${DEV_HOST}:3000` and
    writes the result to `manifest.dev.xml` (gitignored). Use `DEV_HOST=localhost` for direct localhost testing.
  - Vite dev build (`npm run build:dev` / `npm run start`) → copies `manifest.dev.xml` to `dist/manifest.xml`. **Dev
    builds fail loudly if `manifest.dev.xml` is missing** (run `npm run dev:setup` first).
  - Vite prod build (`npm run build`, used inside the Docker image) → copies `manifest.xml` as-is.

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
  `npx -p office-addin-manifest@latest office-addin-manifest validate manifest.xml`.

## Common commands

```bash
npm run build:dev          # vite dev build
npm run build              # vite production build
npm run watch              # vite build --watch
npm run typecheck          # vue-tsc --noEmit
npm run lint               # eslint src
npm run start              # vite dev server (HTTPS, port 3000) — for sideloading
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
- **Don't add a ComposeView/ReadView/RetycLink scanner.** The original skeleton had stubs for those features; out of
  scope.
- **Manifest UUID** is currently the placeholder `00000000-0000-0000-0000-000000000001`. **Replace with a real GUID
  before publishing to AppSource.**
- **Don't re-introduce runtime URL substitution in Docker.** We had a `BASE_URL` / `LANDING_URL` entrypoint script that
  rewrote the manifest and nginx config at container start to keep one image for all envs. It was removed in favour of
  per-environment builds (separate manifest files + `VITE_RETYC_API_URL` build arg). Adding it back means juggling two
  sources of truth for env config — if a new env is needed, build a new image.
