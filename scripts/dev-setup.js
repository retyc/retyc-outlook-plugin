#!/usr/bin/env node
/* eslint-disable */
// One-shot dev setup: regenerates the mkcert leaf certificate AND manifest.dev.xml
// for the given DEV_HOST (the IP/hostname the Windows VM uses to reach this host).
//
// Default: 192.168.122.1 (libvirt default bridge gateway).
// Override with: DEV_HOST=10.0.2.2 npm run dev:setup
//
// Both outputs go to gitignored locations:
//   .certs/dev.pem + .certs/dev-key.pem   (used by webpack-dev-server)
//   manifest.dev.xml                      (sideload this in Outlook)
//
// Re-run this script every time the VM's networking changes — the cert SAN and
// the manifest URLs MUST agree, otherwise Outlook will silently fail to load
// runtimes / icons.

const { spawnSync } = require('child_process')
const fs = require('fs')
const path = require('path')

const DEV_HOST = process.env.DEV_HOST || '192.168.122.1'
const SOURCE_URL = 'https://outlook.retyc.com'   // matches the placeholder in the canonical manifest.xml
const root = path.resolve(__dirname, '..')
const certsDir = path.join(root, '.certs')
const certPath = path.join(certsDir, 'dev.pem')
const keyPath = path.join(certsDir, 'dev-key.pem')

function fail(message) {
  console.error(`error: ${message}`)
  process.exit(1)
}

// 1. mkcert binary present?
const ver = spawnSync('mkcert', ['-version'], { encoding: 'utf8' })
if (ver.error || ver.status !== 0) {
  fail('mkcert is not installed or not on PATH. See https://github.com/FiloSottile/mkcert')
}

// 2. Local CA present? (require an explicit "mkcert -install" first — that step usually needs sudo)
const caRoot = spawnSync('mkcert', ['-CAROOT'], { encoding: 'utf8' })
if (caRoot.status !== 0) {
  fail('Failed to query mkcert -CAROOT.')
}
const caRootPath = caRoot.stdout.trim()
const rootCaPath = path.join(caRootPath, 'rootCA.pem')
if (!fs.existsSync(rootCaPath)) {
  fail(`No local CA at ${rootCaPath}. Run "mkcert -install" first (may require sudo).`)
}

// 3. Wipe stale cert files so the .certs/ folder always reflects the current DEV_HOST.
fs.mkdirSync(certsDir, { recursive: true })
for (const f of fs.readdirSync(certsDir)) {
  if (f.endsWith('.pem')) fs.unlinkSync(path.join(certsDir, f))
}

// 4. Generate a leaf cert that covers DEV_HOST + localhost (so the dev server can be hit by both).
const gen = spawnSync(
  'mkcert',
  ['-cert-file', certPath, '-key-file', keyPath, DEV_HOST, 'localhost', '127.0.0.1', '::1'],
  { stdio: ['ignore', 'inherit', 'inherit'] },
)
if (gen.status !== 0) {
  fail('mkcert failed to issue the leaf certificate.')
}

// 5. Generate manifest.dev.xml with the same DEV_HOST. The canonical manifest.xml has
//    `https://outlook.retyc.com` as the placeholder URL; we rewrite it to `https://<DEV_HOST>:3000`
//    so Outlook in the VM fetches assets from the dev server on the host.
const srcPath = path.join(root, 'manifest.xml')
const outPath = path.join(root, 'manifest.dev.xml')
const TARGET_URL = `https://${DEV_HOST}:3000`
const xml = fs.readFileSync(srcPath, 'utf8').replace(new RegExp(SOURCE_URL.replace(/\./g, '\\.'), 'g'), TARGET_URL)
fs.writeFileSync(outPath, xml)

// 6. Recap.
console.log()
console.log(`✓ Leaf cert  : ${path.relative(root, certPath)}`)
console.log(`               valid for: ${DEV_HOST}, localhost, 127.0.0.1, ::1`)
console.log(`✓ Manifest   : ${path.relative(root, outPath)}`)
console.log(`               URLs target: ${TARGET_URL}`)
console.log()
console.log(`Next steps:`)
console.log(`  1. Import in the Windows VM (once, or whenever you re-run "mkcert -install"):`)
console.log(`       ${rootCaPath}`)
console.log(`     → Trusted Root Certification Authorities (Local Machine)`)
console.log(`  2. Sideload ${path.relative(root, outPath)} in Outlook (Get Add-ins → My add-ins → Add from File).`)
console.log(`  3. npm run start`)
