# Local development with a Windows VM

This add-in runs in **Outlook desktop on Windows** while the dev server runs on
a Linux host. The VM reaches the host on a single IP — defaults below — and
both the leaf cert SAN and the URLs in `manifest.dev.xml` MUST use the same
host. `npm run dev:setup` keeps them in sync.

| QEMU networking                   | Host IP from the VM | When to use                                                        |
|-----------------------------------|---------------------|--------------------------------------------------------------------|
| libvirt default bridge (`virbr0`) | **`192.168.122.1`** | typical `virt-manager` / `qemu-system-x86_64 -netdev bridge` setup |
| user-mode NAT (`-netdev user`)    | **`10.0.2.2`**      | bare QEMU CLI without a tap device                                 |
| custom bridge                     | host LAN IP         | manual `bridge=br0` setups                                         |

If unsure, run `ip route` inside the VM — the gateway IP is the host.

> **Why XML and not JSON?** Classic Outlook desktop on Windows only accepts the
> legacy XML manifest schema (`<OfficeApp xsi:type="MailApp">`). The unified
> Microsoft 365 JSON manifest is supported on new Outlook / Outlook on the web /
> Mac, but not by Classic Outlook's "Add from File" picker.

## One-time setup

### 1. Install the mkcert local CA on the host

```bash
mkcert -install      # may prompt for sudo — registers the local CA in the system trust store
mkcert -CAROOT       # prints e.g. /home/<you>/.local/share/mkcert  (path to rootCA.pem)
```

You only run this once per machine. Re-running is harmless.

### 2. Generate the leaf cert + dev manifest

```bash
npm run dev:setup                                  # uses DEV_HOST=192.168.122.1 by default
DEV_HOST=10.0.2.2 npm run dev:setup                # override for QEMU user-mode
DEV_HOST=192.168.1.42 npm run dev:setup            # override for a LAN-bridged VM
```

This single command writes:

- `.certs/dev.pem` + `.certs/dev-key.pem` — leaf certificate covering `DEV_HOST`, `localhost`, `127.0.0.1`, `::1`
  (consumed by webpack-dev-server).
- `manifest.dev.xml` — the manifest with all URLs rewritten to `https://DEV_HOST:3000`.

**Always re-run after changing the VM's networking.** Otherwise the cert SAN no longer matches the manifest's URLs and
Outlook silently fails to load runtimes/icons.

### 3. Trust the local CA inside the VM

Copy `rootCA.pem` from the host (path printed by `mkcert -CAROOT`) to the VM, then:

1. Rename the file to `rootCA.crt` (Windows associates `.crt` with the cert UI).
2. Double-click → **Install Certificate** → **Local Machine** → **Place all certificates in the following store** → *
   *Browse → Trusted Root Certification Authorities** → **OK**.
    - **Local Machine** is critical — Office processes don't see Current User certs.
3. Verify in Edge inside the VM: `https://192.168.122.1:3000/assets/icon-80.png` must load **without** a security
   warning.
4. If Edge still warns: open `certlm.msc` (Local Machine cert store), confirm the cert is under **Trusted Root
   Certification Authorities → Certificates**.

> Once the rootCA is trusted, regenerating leaf certs (steps 2 above) does NOT require re-importing anything in the VM —
> they're auto-trusted by virtue of being signed by the same rootCA. Only changing the rootCA itself (
`mkcert -uninstall && mkcert -install`) forces a re-import.

### 4. Sideload the manifest in Outlook (VM)

1. Copy `manifest.dev.xml` from the host into the VM (shared folder, scp, USB).
2. In Outlook desktop:
    - Open or compose a message.
    - **Home → Get Add-ins** (or the `…` overflow menu in the message header).
    - **My add-ins** (left panel) → at the bottom: **Add a custom add-in → Add from File**.
    - Pick `manifest.dev.xml` (the file picker filters for `.xml`).
3. Restart Outlook (forced by the dialog).
4. The **Retyc** button appears on the default ribbon, both in compose and read mode.

> "Add from File" greyed out → your account doesn't allow custom add-ins. Use
> a [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) tenant for free testing
> without admin policy roadblocks.
>
> Personal Microsoft accounts (Outlook.com / Hotmail) only support the task pane and ribbon — the `messageSending`
> interception is silently ignored. Test the full flow with an M365 account.

## Daily workflow

```bash
# Host — Linux
npm run start                       # serves https://<DEV_HOST>:3000 using the trusted dev cert
npm run watch                       # alternative: rebuild on change without serving
```

Inside the VM, every change to `src/**` is picked up by webpack-dev-server and
reloaded into the Outlook iframes. No need to re-sideload as long as the
manifest URLs and file names don't change.

If you change `manifest.xml` **or** the VM's networking, run `npm run dev:setup`
again, copy `manifest.dev.xml` to the VM, and re-add the add-in (Outlook caches
the manifest).

## Debugging from the VM

- **Task pane console**: right-click inside the pane → **Inspect**. Opens a WebView2 DevTools window.
- **Dialog console**: same — right-click in the popup once it's visible.
- **Commands runtime** (the headless one that handles `OnMessageSend`): not directly inspectable. Practical workaround —
  log via `console.error`, then open the task pane DevTools where commands runtime errors typically surface in the same
  WebView2 process. Or watch `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\Outlook\<id>\` for runtime traces.
- **Manifest validation**: `npx -p office-addin-manifest@1.13.6 office-addin-manifest validate manifest.xml` on the
  host. (Pin to 1.x — v2.x ESM-requires Node 20+.)
- **Manual reload of the add-in inside Outlook**: disable & re-enable from **Get Add-ins → My add-ins**, or remove +
  re-add the custom add-in.

## Networking troubleshooting

| Symptom                                               | Likely cause                                                 | Fix                                                                     |
|-------------------------------------------------------|--------------------------------------------------------------|-------------------------------------------------------------------------|
| `https://<DEV_HOST>:3000` times out from VM           | Dev server bound to `127.0.0.1` only                         | Confirm `host: '0.0.0.0'` in `webpack.config.js`                        |
| `Invalid Host header` 403                             | webpack-dev-server 5+ host check                             | `allowedHosts: 'all'` already set                                       |
| Cert warning in Edge inside VM                        | rootCA missing from Local Machine store                      | Re-import `rootCA.pem` (Local Machine, not Current User)                |
| Cert warning ONLY for the new IP                      | Leaf cert SAN doesn't include the new IP                     | Re-run `npm run dev:setup` with the right `DEV_HOST`                    |
| Icons missing on the ribbon, "Retyc travaille…" hangs | Outlook can't fetch from `<DEV_HOST>:3000` (cert or network) | Test in Edge first; restart Outlook after each cert change              |
| `Add-in could not start` in Outlook                   | Manifest URL unreachable                                     | Ping the host IP from the VM, check Linux firewall (`sudo iptables -L`) |
| `messageSending` never fires                          | Personal account or Mailbox <1.12                            | Use an M365 account with Outlook ≥ 2024                                 |

## Files involved

```
.certs/dev.pem .certs/dev-key.pem     mkcert leaf for current DEV_HOST (gitignored)
assets/icon-{16,32,48,64,80,128}.png  PNG icons (XML manifest does not accept SVG)
manifest.xml                          canonical manifest, targets localhost
manifest.dev.xml                      generated for VM sideload, targets DEV_HOST (gitignored)
scripts/dev-setup.js                  regenerates cert + manifest from DEV_HOST in one shot
webpack.config.js                     reads .certs/dev.{pem,-key.pem}, binds 0.0.0.0
```
