<p align="center">
  <img width="128" src="https://raw.githubusercontent.com/retyc/retyc-outlook-plugin/master/assets/icon-128.png" alt="Retyc logo" />
</p>

<h1 align="center">Retyc for Outlook</h1>

<p align="center">
  Send large files securely from Outlook — end-to-end encrypted, GDPR-compliant.<br/>
  Files are encrypted client-side in the task pane and replaced by a secure download link in your message body.
</p>

<p align="center">
  <a href="https://github.com/retyc/retyc-outlook-plugin/actions/workflows/main.yml">
    <img src="https://github.com/retyc/retyc-outlook-plugin/actions/workflows/main.yml/badge.svg" alt="CI" />
  </a>
  <img src="https://img.shields.io/badge/outlook-Mailbox%20%E2%89%A5%201.7-blue" alt="Mailbox API ≥ 1.7" />
  <a href="https://github.com/retyc/retyc-outlook-plugin/blob/master/LICENSE">
    <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License: MIT" />
  </a>
</p>

---

## Features

- **Drag-and-drop into the task pane** — files never go through Outlook's attachment system, so they never traverse
  Exchange in cleartext
- **End-to-end encryption** — encryption happens client-side in the task pane via
  the [Retyc SDK](https://www.npmjs.com/package/@retyc/sdk) (age / X25519 hybrid post-quantum) before bytes leave the
  WebView2 process
- **Passphrase support** — required for recipients without a Retyc account, or for passphrase-only transfers without a
  recipient list (≥ 8 characters)
- **OIDC Device Flow auth** — log in with your Retyc account directly from Outlook
- **Configurable transfer expiry** — pick how long the transfer link stays alive, from the task pane Options
- **Bilingual UI** — auto-detects Outlook's display language; English / French bundled, switchable from the Account tab
- **Clean emails** — a Retyc download link is appended to the message body; recipients are written into Outlook's `To`
  field on confirm

## Requirements

- An Outlook client supporting **Mailbox API 1.7+**: Classic Outlook on Windows, new Outlook on Windows, Outlook on Mac,
  Outlook on the web
- A [Retyc](https://retyc.com) account

## Usage

### 1. Log in

Click the **Send via Retyc** button in the compose ribbon to open the task pane, then click **Log in with Retyc**. A
code and URL appear — open the URL in your browser, enter the code, and authenticate. You can pin the task pane (📌 in
the pane's header) to keep it open across drafts.

### 2. Send files via Retyc

In a compose window with the task pane open:

1. **Drop files** into the dropzone (or click to pick).
2. The recipient list mirrors Outlook's `To` / `Cc` / `Bcc` fields — edit them in Outlook and the pane stays in sync.
   Your own address is filtered out automatically.
3. Optionally toggle **Use a passphrase** in the Options panel (≥ 8 chars) — required if you have no recipients, or for
   recipients without a Retyc account.
4. Pick a transfer expiry from the same Options panel.
5. Click **Encrypt & insert Retyc link** — the pane uploads the encrypted bytes, writes the merged recipient list into
   Outlook's `To` field, and appends a Retyc download link to the message body.
6. Click **Send** in Outlook normally.

The Account tab shows your current quota (storage + transfers) and lets you switch the UI language.

## Development

### Prerequisites

- Node.js 24+ (declared in `package.json` engines and `.nvmrc`)
- The `@retyc/sdk` package (declared as a runtime dependency, fetched from the public npm registry)
- An Outlook account that supports Mailbox API 1.7+

### Setup

```bash
npm install
npm run build:dev      # vite development build
npm run build          # vite production build
npm run watch          # rebuild on file changes
npm run typecheck      # vue-tsc --noEmit
npm run lint           # eslint src
npm run start          # vite dev server (HTTPS, port 3000)
npm run dev:setup      # regenerate mkcert leaf cert + manifest.dev.xml (used by DEV_VM workflow)
```

### Sideloading the add-in

1. Run `npm run start` to serve the bundle over HTTPS at `https://localhost:3000`.
2. Sideload `manifest.xml` into Outlook:
    - **Classic Outlook on Windows**: ribbon **Home → Get Add-ins → My add-ins → Add a custom add-in → Add from File**,
      pick `manifest.xml`.
    - **Outlook on the web / new Outlook on Windows / Mac**: Settings → Mail → Customize actions → Get Add-ins → My
      add-ins → Add a custom add-in → From File.
    - **Microsoft 365 admin (org-wide)**: Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps.
3. Open a compose window — the **Send via Retyc** button appears in the ribbon (group label "Retyc"). Click it to open
   the task pane.

For development against a Windows VM (recommended for testing Classic Outlook desktop on Linux hosts), see
[`DEV_VM.md`](DEV_VM.md) — it covers the mkcert setup, the IP-rewritten `manifest.dev.xml`, and the sideload procedure.

### Production hosting

A multi-stage `Dockerfile` is provided. It builds the static bundle with Node 24 and serves it through nginx 1.30
alpine-slim (~13 MB image). Images are **per-environment**: the manifest baked in (`manifest.xml`) and the API URL
(`VITE_RETYC_API_URL`, defaults to `https://api.retyc.com`) are resolved at build time. Build a separate image for each
environment.

```bash
# Prod
docker build -t retyc-outlook-plugin:prod .

# Other environments — provide your own gitignored manifest, override the API URL at build time
cp manifest.<env>.xml manifest.xml
docker build \
  --build-arg VITE_RETYC_API_URL=https://api.<env>.example.com \
  -t retyc-outlook-plugin:<env> .
```

The `/` redirect is hardcoded to `/taskpane.html` (relative). Browsers hitting the bare-domain root land on the task
pane directly — Outlook itself fetches `/taskpane.html` and never relies on the redirect.

CI/release workflows publish to **ghcr.io/retyc/retyc-outlook-plugin** (and Docker Hub `retyc/retyc-outlook-plugin` once
the credentials block in `_docker.yml` is uncommented).

### Project structure

```
src/
  main.ts         # bootstrap (Buffer polyfill, lucide icons, i18n, Office.onReady → mount)
  App.vue         # auth gate + UTabs (Transfer / Account)
  components/     # LoginWall, TransferTab, AccountTab, DropZone (lazy-loaded .vue)
  composables/    # useAuth, useTransfer, useDropOverlay
  shared/         # constants, sdk-factory, storage, token-store, upload pipeline
  i18n.ts         # vue-i18n singleton + locale detection
  locales/        # en.ts, fr.ts message bundles
  assets/         # custom.css with the @theme block
public/           # PNG icons (XML manifest does not accept SVG)
scripts/          # dev-setup.js (regenerates mkcert leaf + manifest.dev.xml)
taskpane.html     # Vite HTML entry (kept at repo root)
manifest.xml      # Office Add-in XML manifest (Outlook MailApp, VersionOverrides V1_0 + V1_1)
manifest.dev.xml  # generated for VM sideload, gitignored
Dockerfile        # multi-stage build → nginx:1.30.0-alpine-slim, ~13 MB
```

### Architecture

The add-in declares **a single runtime** — the task pane — and does all its work there. Persistence is delegated to
`Office.context.roamingSettings` (mailbox-scoped, syncs across devices) wrapped by a small abstraction that falls back
to `localStorage` when the page is opened standalone in a browser.

**Why no `OnMessageSend` interception?** Office.js's event runtime is isolated from the compose runtime on Classic
Outlook desktop — `getAttachmentContentAsync` returns empty content for attachments unless the draft has been saved to
Exchange first. Forcing a `saveAsync` would push the unencrypted attachment to Exchange in cleartext, breaking the E2EE
threat model. We sidestep the whole problem by **never using Outlook's attachment system**: files come from a dropzone
in the task pane and go through `file.arrayBuffer()` → SDK encryption → upload.

**Send-via-Retyc flow**

```
User opens a compose window and pins/opens the Retyc task pane
  └─ User drops files into the task pane (HTML5 drag-drop or file picker)
  └─ User types recipients (textarea, comma/semicolon-separated)
  └─ User optionally types a passphrase (≥ 8 chars)
  └─ User clicks "Encrypt & insert Retyc link"
       ├─ Read each File via file.arrayBuffer() — bytes stay in the WebView2 process
       ├─ @retyc/sdk encrypts client-side (age / X25519 hybrid PQ)
       ├─ SDK uploads encrypted chunks to the Retyc API
       ├─ item.to.setAsync(recipients) — only if recipients were provided
       └─ body.setAsync(existingHtml + retycLinkSnippet)
  └─ User clicks Send in Outlook → email goes through Exchange with no attachments,
     just the Retyc link in the body
```

## License

[MIT](LICENSE) — © Retyc / TripleStack SAS
