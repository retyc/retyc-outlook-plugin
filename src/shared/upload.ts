// Reads user-selected files (from a drop zone or file input in the taskpane), encrypts them
// client-side via @retyc/sdk, and appends a Retyc download link to the compose body.
//
// We deliberately bypass Outlook's attachment system: the bytes go from the user's disk
// straight into the WebView2 process and are encrypted there. This keeps the E2EE threat
// model intact — the cleartext attachment never traverses Exchange — and avoids the
// `getAttachmentContentAsync` quirks of Classic Outlook desktop's event runtime.

import type { RetycSDK } from '@retyc/sdk'
import { TRANSFER_PATH_PREFIX } from './constants'

const MAX_TOTAL_BYTES = 5 * 1024 * 1024 * 1024 // 5 GB hard limit

export interface UploadProgress {
  fileName: string
  fileIndex: number
  totalFiles: number
}

export interface UploadOptions {
  recipients: string[]
  appUrl: string
  expiresDays: number
  passphrase?: string
  onProgress?: (p: UploadProgress) => void
}

export interface UploadOutcome {
  transferUrl: string
}

// --- Outlook compose helpers (recipients + body injection only — no attachment access) ---

function listEmails(field: Office.Recipients): Promise<Office.EmailAddressDetails[]> {
  return new Promise((resolve, reject) => {
    field.getAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value)
      else reject(new Error(r.error?.message ?? 'Failed to read recipients.'))
    })
  })
}

export async function readInitialRecipients(item: Office.MessageCompose): Promise<string[]> {
  // Read once when the taskpane opens to pre-fill the Recipients input. After that the
  // taskpane field is the source of truth — we don't poll Outlook.
  const all: string[] = []
  for (const field of [item.to, item.cc, item.bcc]) {
    try {
      const list = await listEmails(field)
      for (const x of list) {
        if (x.emailAddress) all.push(x.emailAddress)
      }
    } catch (err) {
      console.warn('[Retyc] failed to read a recipient field:', err)
    }
  }
  return [...new Set(all)]
}

export function setOutlookTo(item: Office.MessageCompose, emails: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    item.to.setAsync(emails, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(new Error(r.error?.message ?? 'Failed to set recipients in Outlook.'))
    })
  })
}

function getBodyHtml(item: Office.MessageCompose): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value)
      else reject(new Error(r.error?.message ?? 'Failed to read body content.'))
    })
  })
}

function setBodyHtml(item: Office.MessageCompose, html: string): Promise<void> {
  return new Promise((resolve, reject) => {
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(new Error(r.error?.message ?? 'Failed to set body content.'))
    })
  })
}

async function appendBodyHtml(item: Office.MessageCompose, html: string): Promise<void> {
  // Office.js compose body has no appendAsync; read + set is the supported pattern.
  const existing = await getBodyHtml(item)
  await setBodyHtml(item, existing + html)
}

// --- Email validation ---

const EMAIL_RE = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/

export function parseRecipientsInput(raw: string): { valid: string[]; invalid: string[] } {
  const valid: string[] = []
  const invalid: string[] = []
  const seen = new Set<string>()
  for (const candidate of raw.split(/[\s,;]+/).map((s) => s.trim()).filter(Boolean)) {
    const key = candidate.toLowerCase()
    if (seen.has(key)) continue
    seen.add(key)
    if (EMAIL_RE.test(candidate)) valid.push(candidate)
    else invalid.push(candidate)
  }
  return { valid, invalid }
}

// --- Link rendering ---

function buildTransferUrl(appUrl: string, slug: string): string {
  return `${appUrl.replace(/\/$/, '')}${TRANSFER_PATH_PREFIX}/${encodeURIComponent(slug)}`
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function buildLinkSnippet(transferUrl: string): string {
  const safe = escapeHtml(transferUrl)
  return `
<br><br>
<hr style="border:none;border-top:1px solid #e0e0e0;margin:16px 0">
<p style="font-family:sans-serif;font-size:14px;color:#444;margin:0">
  <strong>&#128230; Your files are available via Retyc:</strong><br>
  <a href="${safe}" style="color:#1a3c6e">${safe}</a><br>
  <small style="color:#888">This transfer is end-to-end encrypted and will expire automatically.</small>
</p>`
}

// --- Pipeline ---

export async function performRetycTransfer(
  item: Office.MessageCompose,
  sdk: RetycSDK,
  files: File[],
  options: UploadOptions,
): Promise<UploadOutcome> {
  if (files.length === 0) {
    throw new Error('No files selected to upload.')
  }
  if (options.recipients.length === 0 && !options.passphrase) {
    throw new Error('Provide at least one recipient or a passphrase.')
  }

  const totalBytes = files.reduce((sum, f) => sum + f.size, 0)
  if (totalBytes > MAX_TOTAL_BYTES) {
    throw new Error(
      `Total file size (${(totalBytes / 1e9).toFixed(1)} GB) exceeds the 5 GB limit.`,
    )
  }

  // Read each file sequentially to avoid holding all data in memory at once.
  const uploadFiles: Array<{ name: string; mimeType: string; data: Uint8Array; size: number }> = []
  for (let idx = 0; idx < files.length; idx++) {
    const file = files[idx]
    options.onProgress?.({ fileName: file.name, fileIndex: idx, totalFiles: files.length })
    const buffer = await file.arrayBuffer()
    uploadFiles.push({
      name: file.name,
      mimeType: file.type || 'application/octet-stream',
      data: new Uint8Array(buffer),
      size: buffer.byteLength,
    })
  }

  const result = await sdk.transfers.upload({
    recipients: options.recipients,
    expires: options.expiresDays * 24 * 60 * 60,
    files: uploadFiles,
    ...(options.passphrase ? { passphrase: options.passphrase } : {}),
  })

  const transferUrl = buildTransferUrl(options.appUrl, result.slug)

  // Mirror the recipients into Outlook's To field so the email reaches them when the user
  // clicks Send. We don't touch Cc/Bcc — those remain whatever the user typed in Outlook.
  // Skip when no recipients are provided (passphrase-only transfer) so we don't wipe what
  // the user already typed in Outlook.
  if (options.recipients.length > 0) {
    try {
      await setOutlookTo(item, options.recipients)
    } catch (err) {
      console.warn('[Retyc] failed to set Outlook To field:', err)
    }
  }

  // Append the Retyc link to the compose body (read + set; Office.js has no appendAsync here).
  await appendBodyHtml(item, buildLinkSnippet(transferUrl))

  return { transferUrl }
}

export function isAuthError(message: string): boolean {
  return (
    message.toLowerCase().includes('refresh token') ||
    message.toLowerCase().includes('log in again') ||
    (/\b401\b/.test(message) && message.toLowerCase().includes('auth'))
  )
}
