// Reads user-selected files (from a drop zone or file input in the taskpane), encrypts them
// client-side via @retyc/sdk, and appends a Retyc download link to the compose body.
//
// We deliberately bypass Outlook's attachment system: the bytes go from the user's disk
// straight into the WebView2 process and are encrypted there. This keeps the E2EE threat
// model intact — the cleartext attachment never traverses Exchange — and avoids the
// `getAttachmentContentAsync` quirks of Classic Outlook desktop's event runtime.

import type { RetycSDK, UploadProgress as SdkUploadProgress } from '@retyc/sdk'

const MAX_TOTAL_BYTES = 5 * 1024 * 1024 * 1024 // 5 GB hard limit

export type UploadPhase = 'reading' | 'uploading'

export interface UploadProgress {
  phase: UploadPhase
  fileName: string
  fileIndex: number   // 0-based
  totalFiles: number
  uploadedBytes: number
  totalBytes: number
  ratio: number       // 0..1
}

export interface UploadOptions {
  recipients: string[]
  expires: number | null  // seconds; null = never expires
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

function setOutlookTo(item: Office.MessageCompose, emails: string[]): Promise<void> {
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

// Inserts the snippet immediately before Outlook's signature block (id="Signature" on Classic
// desktop, id="signature" on some web/new clients), so the link sits inside the user's message
// rather than dangling under the signature. Falls back to appending if no signature is detected.
function insertBeforeSignature(existing: string, snippet: string): string {
  const match = existing.match(/<div\b[^>]*\bid=["']signature["'][^>]*>/i)
  if (match && match.index !== undefined) {
    return existing.slice(0, match.index) + snippet + existing.slice(match.index)
  }
  return existing + snippet
}

async function injectLinkSnippet(item: Office.MessageCompose, snippet: string): Promise<void> {
  // Office.js compose body has no appendAsync; read + set is the supported pattern.
  const existing = await getBodyHtml(item)
  await setBodyHtml(item, insertBeforeSignature(existing, snippet))
}

// --- Link rendering ---

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function formatExpiry(seconds: number): string {
  if (seconds >= 2 * 31536000) return `${Math.round(seconds / 31536000)} years`
  if (seconds >= 31536000)     return '1 year'
  if (seconds >= 2 * 2592000)  return `${Math.round(seconds / 2592000)} months`
  if (seconds >= 2592000)      return '1 month'
  if (seconds >= 2 * 86400)    return `${Math.round(seconds / 86400)} days`
  if (seconds >= 86400)        return '1 day'
  if (seconds >= 2 * 3600)     return `${Math.round(seconds / 3600)} hours`
  if (seconds >= 3600)         return '1 hour'
  return `${Math.round(seconds / 60)} minutes`
}

function buildLinkSnippet(transferUrl: string, expires: number | null): string {
  const safe = escapeHtml(transferUrl)
  const expiryLine = expires !== null
    ? ` &nbsp;&bull;&nbsp; This link expires in ${escapeHtml(formatExpiry(expires))}`
    : ''
  return `
<br><br>
<div style="font-family:sans-serif;font-size:14px;color:#444;border:1px solid #e0e0e0;border-radius:6px;padding:14px 16px;display:inline-block;max-width:560px">
  <strong>&#128230; Your files are available via Retyc:</strong><br>
  <a href="${safe}" style="color:#1a3c6e">${safe}</a><br>
  <small style="color:#888">&#128274; End-to-end encrypted${expiryLine}</small>
</div>`
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
    options.onProgress?.({
      phase: 'reading',
      fileName: file.name,
      fileIndex: idx,
      totalFiles: files.length,
      uploadedBytes: 0,
      totalBytes,
      ratio: 0,
    })
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
    expires: options.expires as unknown as number, // null = never expires; SDK types are conservative but backend accepts it
    files: uploadFiles,
    ...(options.passphrase ? { passphrase: options.passphrase } : {}),
    ...(options.onProgress
      ? {
          onProgress: (p: SdkUploadProgress) => {
            options.onProgress!({
              phase: 'uploading',
              fileName: p.currentFile.name,
              fileIndex: p.currentFile.index,
              totalFiles: p.currentFile.total,
              uploadedBytes: p.uploadedBytes,
              totalBytes: p.totalBytes,
              ratio: p.ratio,
            })
          },
        }
      : {}),
  })

  const transferUrl = result.webUrl

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

  // Inject the Retyc link into the compose body, before Outlook's signature when present.
  await injectLinkSnippet(item, buildLinkSnippet(transferUrl, options.expires))

  return { transferUrl }
}

export function isAuthError(message: string): boolean {
  return (
    message.toLowerCase().includes('refresh token') ||
    message.toLowerCase().includes('log in again') ||
    (/\b401\b/.test(message) && message.toLowerCase().includes('auth'))
  )
}
