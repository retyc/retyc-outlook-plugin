import './taskpane.css'
import { getSDK, invalidateSDK } from '../shared/sdk-factory'
import {
  type RetycSettings,
  getAllSettings,
  setAllSettings,
} from '../shared/settings'
import type { UserInfo } from '../shared/messages'
import {
  isAuthError,
  parseRecipientsInput,
  performRetycTransfer,
  readInitialRecipients,
} from '../shared/upload'

// Cached user info — fetched once after login, cleared on logout or API URL change.
let _cachedUserInfo: UserInfo | null = null

// Files the user has dropped/picked, awaiting encryption.
let _selectedFiles: File[] = []

// True once the user has typed in the Recipients input — beyond that point we never overwrite
// it with what's in Outlook. The "Use Outlook recipients" button explicitly forces a refill.
let _recipientsTouched = false

function showAuth(id: string): void {
  document.querySelectorAll<HTMLElement>('.auth-block').forEach((el) => el.classList.add('hidden'))
  document.getElementById(id)?.classList.remove('hidden')
}

function showFeedback(message: string, type: 'success' | 'error'): void {
  const el = document.getElementById('save-feedback')!
  el.textContent = message
  el.className = `feedback ${type}`
  el.classList.remove('hidden')
  setTimeout(() => el.classList.add('hidden'), 5000)
}

let pollTimer: ReturnType<typeof setTimeout> | null = null
function stopPolling(): void {
  if (pollTimer !== null) {
    clearTimeout(pollTimer)
    pollTimer = null
  }
}

function formatExpiry(expiresAt: number): string {
  const seconds = expiresAt - Math.floor(Date.now() / 1000)
  if (seconds <= 0) return 'Token expired'
  if (seconds < 60) return `Expires in ${seconds}s`
  const minutes = Math.floor(seconds / 60)
  if (minutes < 60) return `Expires in ${minutes}m`
  const hours = Math.floor(minutes / 60)
  const mins = minutes % 60
  if (hours < 24) return mins > 0 ? `Expires in ${hours}h ${mins}m` : `Expires in ${hours}h`
  const days = Math.floor(hours / 24)
  return `Expires in ${days}d`
}

interface AuthSnapshot {
  authenticated: boolean
  userInfo?: UserInfo
  tokenExpiresAt?: number
}

async function readAuthSnapshot(): Promise<AuthSnapshot> {
  try {
    const sdk = await getSDK()
    const tokens = await sdk.auth.getTokens()
    if (!tokens) return { authenticated: false }

    if (!_cachedUserInfo) {
      try {
        const me = await sdk.user.getMe()
        _cachedUserInfo = {
          fullName: me.user.full_name,
          email: me.user.email,
        }
      } catch { /* user info is optional */ }
    }

    return {
      authenticated: true,
      userInfo: _cachedUserInfo ?? undefined,
      tokenExpiresAt: tokens.expiresAt,
    }
  } catch {
    return { authenticated: false }
  }
}

function renderAuthenticated(snapshot: AuthSnapshot): void {
  const fullNameEl = document.getElementById('user-fullname')!
  const emailEl = document.getElementById('user-email')!
  const expiresEl = document.getElementById('token-expires')!

  if (snapshot.userInfo) {
    fullNameEl.textContent = snapshot.userInfo.fullName ?? ''
    fullNameEl.style.display = snapshot.userInfo.fullName ? '' : 'none'
    emailEl.textContent = snapshot.userInfo.email
  } else {
    fullNameEl.style.display = 'none'
    emailEl.textContent = ''
  }

  expiresEl.textContent = snapshot.tokenExpiresAt ? formatExpiry(snapshot.tokenExpiresAt) : ''
  showAuth('auth-authenticated')
}

function fillSettings(s: RetycSettings): void {
  ;(document.getElementById('api-url') as HTMLInputElement).value = s.apiUrl
  ;(document.getElementById('app-url') as HTMLInputElement).value = s.appUrl
  ;(document.getElementById('expires-days') as HTMLInputElement).value = String(s.expiresDays)
}

function setSectionCollapsed(id: string, collapsed: boolean): void {
  document.getElementById(id)?.classList.toggle('collapsed', collapsed)
}

async function loadStatus(): Promise<void> {
  showAuth('auth-loading')
  fillSettings(getAllSettings())
  const snapshot = await readAuthSnapshot()
  if (snapshot.authenticated) renderAuthenticated(snapshot)
  else showAuth('auth-unauthenticated')

  // Authentication card collapsed by default once the user is signed in (compose is the focus then).
  // API & App is always collapsed by default — most users won't change those settings.
  setSectionCollapsed('auth-section', snapshot.authenticated)
  setSectionCollapsed('api-app-section', true)

  void refreshComposeSection(snapshot.authenticated)
}

async function refreshToken(): Promise<void> {
  const btn = document.getElementById('btn-refresh') as HTMLButtonElement
  btn.disabled = true
  const previousLabel = btn.textContent
  btn.textContent = '…'
  try {
    const sdk = await getSDK()
    const tokens = await sdk.auth.refresh()
    document.getElementById('token-expires')!.textContent = formatExpiry(tokens.expiresAt)
  } catch (err) {
    showFeedback(
      `Refresh failed: ${err instanceof Error ? err.message : String(err)}`,
      'error',
    )
  } finally {
    btn.disabled = false
    btn.textContent = previousLabel
  }
}

async function startLogin(): Promise<void> {
  // Persist any pending edits in the settings form so getSDK() reads the latest API URL.
  const validation = validateForm()
  if (validation.error || !validation.settings) {
    showFeedback(validation.error ?? 'Invalid settings.', 'error')
    return
  }
  try {
    await setAllSettings(validation.settings)
    invalidateSDK()
    _cachedUserInfo = null
  } catch (err) {
    showFeedback(
      `Failed to save settings: ${err instanceof Error ? err.message : String(err)}`,
      'error',
    )
    return
  }

  showAuth('auth-device-flow')
  try {
    const sdk = await getSDK()
    const flow = await sdk.auth.startDeviceFlow()

    flow.poll().catch((err: unknown) => {
      console.error('[Retyc] Device flow polling error:', err)
    })

    const urlEl = document.getElementById('verification-url') as HTMLAnchorElement
    urlEl.href = flow.verificationUriComplete ?? flow.verificationUri
    urlEl.textContent = flow.verificationUri
    document.getElementById('user-code')!.textContent = flow.userCode

    const deadline = Date.now() + flow.expiresIn * 1000
    const tick = async (): Promise<void> => {
      if (pollTimer === null) return
      if (Date.now() > deadline) {
        stopPolling()
        showFeedback('Authentication timed out. Please try again.', 'error')
        showAuth('auth-unauthenticated')
        return
      }
      const snapshot = await readAuthSnapshot()
      if (snapshot.authenticated) {
        stopPolling()
        renderAuthenticated(snapshot)
        showFeedback('Successfully logged in to Retyc.', 'success')
        // Once logged in, collapse the auth card so the compose card has the spotlight.
        setSectionCollapsed('auth-section', true)
        void refreshComposeSection(true)
      } else {
        pollTimer = setTimeout(tick, 3000)
      }
    }
    pollTimer = setTimeout(tick, 3000)
  } catch (err) {
    showFeedback(`Login failed: ${err instanceof Error ? err.message : String(err)}`, 'error')
    showAuth('auth-unauthenticated')
  }
}

function cancelLogin(): void {
  stopPolling()
  showAuth('auth-unauthenticated')
}

async function logout(): Promise<void> {
  stopPolling()
  _cachedUserInfo = null
  try {
    const sdk = await getSDK()
    await sdk.auth.logout()
  } catch (err) {
    console.error('[Retyc] Logout error:', err)
  }
  showAuth('auth-unauthenticated')
  // Re-open the auth card so the Log in button is in plain sight.
  setSectionCollapsed('auth-section', false)
  showFeedback('Logged out.', 'success')
  void refreshComposeSection(false)
}

function validateSecureUrl(value: string, label: string): string | null {
  if (!value) return `${label} is required.`
  let url: URL
  try {
    url = new URL(value)
  } catch {
    return `${label} is not a valid URL.`
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') {
    return `${label} must use http or https.`
  }
  return null
}

interface ValidationResult {
  error: string | null
  settings?: RetycSettings
}

function validateForm(): ValidationResult {
  const apiUrl = (document.getElementById('api-url') as HTMLInputElement).value.trim()
  const appUrl = (document.getElementById('app-url') as HTMLInputElement).value.trim()
  const expiresDays = parseInt(
    (document.getElementById('expires-days') as HTMLInputElement).value,
    10,
  )

  const err =
    validateSecureUrl(apiUrl, 'API URL') ??
    validateSecureUrl(appUrl, 'App URL') ??
    (!Number.isFinite(expiresDays) || expiresDays < 1 || expiresDays > 365
      ? 'Expiry must be between 1 and 365 days.'
      : null)

  if (err) return { error: err }
  // autoSend is reserved for a future automation; persist a stable default.
  return { error: null, settings: { apiUrl, appUrl, expiresDays, autoSend: true } }
}

async function saveSettings(e: Event): Promise<void> {
  e.preventDefault()

  const { error, settings } = validateForm()
  if (error || !settings) {
    showFeedback(error ?? 'Invalid settings.', 'error')
    return
  }

  try {
    await setAllSettings(settings)
    invalidateSDK() // The API URL may have changed.
    _cachedUserInfo = null
    showFeedback('Settings saved.', 'success')
  } catch (err) {
    showFeedback(
      `Failed to save settings: ${err instanceof Error ? err.message : String(err)}`,
      'error',
    )
  }
}

// --- Compose section ---

type ComposeView = 'prepare' | 'uploading' | 'done' | 'error'

function getComposeItem(): Office.MessageCompose | null {
  const item = Office.context.mailbox?.item as Office.MessageCompose | undefined
  // `to` is a Recipients object only on the compose-side item (read-side has `to` as
  // EmailAddressDetails[], so the typeof guard distinguishes them).
  if (item && 'to' in item && item.to && typeof item.to.getAsync === 'function') {
    return item
  }
  return null
}

function showComposeView(view: ComposeView): void {
  for (const id of ['compose-prepare', 'compose-uploading', 'compose-done', 'compose-error']) {
    document.getElementById(id)?.classList.add('hidden')
  }
  document.getElementById(`compose-${view}`)?.classList.remove('hidden')
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`
}

function getCurrentRecipients(): { valid: string[]; invalid: string[] } {
  const input = document.getElementById('recipients-input') as HTMLTextAreaElement | null
  if (!input) return { valid: [], invalid: [] }
  return parseRecipientsInput(input.value)
}

function renderRecipientsState(): void {
  const { valid, invalid } = getCurrentRecipients()
  const countEl = document.getElementById('recipients-count')
  const errorEl = document.getElementById('recipients-error')

  if (countEl) {
    if (valid.length === 0 && invalid.length === 0) {
      countEl.textContent = 'No recipients yet.'
      countEl.className = 'muted empty'
    } else if (invalid.length > 0) {
      countEl.textContent = `${valid.length} valid, ${invalid.length} invalid`
      countEl.className = 'muted invalid'
    } else {
      countEl.textContent = `${valid.length} recipient${valid.length === 1 ? '' : 's'}`
      countEl.className = 'muted'
    }
  }

  if (errorEl) {
    if (invalid.length > 0) {
      errorEl.textContent = `Invalid email${invalid.length === 1 ? '' : 's'}: ${invalid.join(', ')}`
      errorEl.classList.remove('hidden')
    } else {
      errorEl.classList.add('hidden')
      errorEl.textContent = ''
    }
  }

  updateSendButtonState()
}

function getCurrentPassphrase(): string {
  const input = document.getElementById('compose-passphrase') as HTMLInputElement | null
  return input ? input.value : ''
}

function updateSendButtonState(): void {
  const sendButton = document.getElementById('btn-send-via-retyc') as HTMLButtonElement | null
  if (!sendButton) return
  const hasFiles = _selectedFiles.length > 0
  const { valid, invalid } = getCurrentRecipients()
  const hasValidRecipients = valid.length > 0 && invalid.length === 0
  const passphrase = getCurrentPassphrase().trim()
  // Allow long-enough passphrases as a substitute for recipients (passphrase-only transfer).
  const hasUsablePassphrase = passphrase.length >= 8

  let disabled = false
  let title = ''
  if (!hasFiles) {
    disabled = true
    title = 'Drop or pick files first.'
  } else if (invalid.length > 0) {
    disabled = true
    title = 'Fix the invalid email addresses first.'
  } else if (!hasValidRecipients && !hasUsablePassphrase) {
    disabled = true
    title = 'Add at least one recipient, or set a passphrase (≥ 8 chars).'
  }
  sendButton.disabled = disabled
  sendButton.title = title
}

function renderSelectedFiles(): void {
  const empty = document.querySelector<HTMLDivElement>('.dropzone-empty')!
  const filesPanel = document.getElementById('dropzone-files')!
  const meta = document.getElementById('file-list-meta')!
  const list = document.getElementById('file-list')!
  const summary = document.getElementById('file-list-summary')!

  if (_selectedFiles.length === 0) {
    empty.classList.remove('hidden')
    filesPanel.classList.add('hidden')
    meta.classList.add('hidden')
    updateSendButtonState()
    return
  }

  empty.classList.add('hidden')
  filesPanel.classList.remove('hidden')
  meta.classList.remove('hidden')

  list.innerHTML = ''
  let total = 0
  for (let i = 0; i < _selectedFiles.length; i++) {
    const file = _selectedFiles[i]
    total += file.size

    const li = document.createElement('li')
    const name = document.createElement('span')
    name.className = 'file-name'
    name.textContent = file.name
    const size = document.createElement('span')
    size.className = 'file-size'
    size.textContent = formatSize(file.size)
    const remove = document.createElement('button')
    remove.type = 'button'
    remove.className = 'file-remove'
    remove.title = 'Remove'
    remove.textContent = '×'
    const indexAtClick = i
    remove.addEventListener('click', (ev) => {
      ev.stopPropagation()
      _selectedFiles.splice(indexAtClick, 1)
      renderSelectedFiles()
    })

    li.appendChild(name)
    li.appendChild(size)
    li.appendChild(remove)
    list.appendChild(li)
  }
  summary.textContent = `${_selectedFiles.length} file${_selectedFiles.length === 1 ? '' : 's'} — ${formatSize(total)}`

  updateSendButtonState()
}

function setRecipientsInputValue(emails: string[]): void {
  const input = document.getElementById('recipients-input') as HTMLTextAreaElement | null
  if (!input) return
  input.value = emails.join(', ')
  renderRecipientsState()
}

async function fillRecipientsFromOutlook(item: Office.MessageCompose, force: boolean): Promise<void> {
  if (!force && _recipientsTouched) return
  try {
    const emails = await readInitialRecipients(item)
    if (emails.length > 0) setRecipientsInputValue(emails)
  } catch (err) {
    console.warn('[Retyc] failed to read Outlook recipients:', err)
  }
}

function addFiles(files: FileList | File[]): void {
  for (const f of Array.from(files)) {
    // Deduplicate by name + size + lastModified to avoid double-adding the same drop.
    const dup = _selectedFiles.some(
      (s) => s.name === f.name && s.size === f.size && s.lastModified === f.lastModified,
    )
    if (!dup) _selectedFiles.push(f)
  }
  renderSelectedFiles()
}

function clearSelectedFiles(): void {
  _selectedFiles = []
  renderSelectedFiles()
}

async function refreshComposeSection(authenticated: boolean): Promise<void> {
  const composeSection = document.getElementById('compose-section')!
  const item = getComposeItem()
  // Hide the entire compose card unless we're (a) actually composing AND (b) signed in.
  if (!item || !authenticated) {
    composeSection.classList.add('hidden')
    return
  }

  composeSection.classList.remove('hidden')
  // Pre-fill the recipients input from Outlook only if the user hasn't typed yet.
  await fillRecipientsFromOutlook(item, false)
  showComposeView('prepare')
  renderSelectedFiles()
  renderRecipientsState()
}

function showComposePassphraseError(message: string): void {
  const el = document.getElementById('compose-passphrase-error')!
  el.textContent = message
  el.classList.remove('hidden')
}

function clearComposePassphraseError(): void {
  document.getElementById('compose-passphrase-error')?.classList.add('hidden')
}

function showComposeError(message: string): void {
  document.getElementById('compose-error-message')!.textContent = message
  showComposeView('error')
}

async function sendViaRetyc(): Promise<void> {
  const item = getComposeItem()
  if (!item) {
    showComposeError('No active compose item.')
    return
  }

  if (_selectedFiles.length === 0) {
    showComposePassphraseError('Drop at least one file in the box above.')
    return
  }

  const { valid: recipients, invalid } = getCurrentRecipients()
  if (invalid.length > 0) {
    showComposePassphraseError(`Fix the invalid email addresses: ${invalid.join(', ')}`)
    return
  }

  // Persist any pending settings edits, same logic as the login button.
  const validation = validateForm()
  if (validation.error || !validation.settings) {
    showComposeError(validation.error ?? 'Invalid settings.')
    return
  }
  try {
    await setAllSettings(validation.settings)
    invalidateSDK()
  } catch (err) {
    showComposeError(`Failed to save settings: ${err instanceof Error ? err.message : String(err)}`)
    return
  }

  const passphraseInput = document.getElementById('compose-passphrase') as HTMLInputElement | null
  const passphraseRaw = passphraseInput?.value.trim() ?? ''
  const passphrase = passphraseRaw || undefined

  // Validate first — only clear the DOM value once we're sure we're proceeding to upload.
  if (passphrase !== undefined && passphrase.length < 8) {
    showComposePassphraseError('Passphrase must be at least 8 characters.')
    passphraseInput?.focus()
    return
  }

  // Need at least recipients OR a passphrase — otherwise nobody can decrypt the transfer.
  if (recipients.length === 0 && passphrase === undefined) {
    showComposePassphraseError('Add at least one recipient, or set a passphrase (≥ 8 characters).')
    return
  }
  clearComposePassphraseError()

  // Now that validation passed, wipe the passphrase from the DOM (it's still held in the
  // local `passphrase` closure variable for the upload itself).
  if (passphraseInput) passphraseInput.value = ''

  // Snapshot the file list so a late × click or drag-drop doesn't mutate the upload mid-flight.
  const filesToUpload = [..._selectedFiles]

  const settings = validation.settings
  const sendButton = document.getElementById('btn-send-via-retyc') as HTMLButtonElement | null
  if (sendButton) sendButton.disabled = true
  showComposeView('uploading')

  const statusEl = document.getElementById('compose-upload-status')
  const fillEl = document.getElementById('compose-progress-fill') as HTMLDivElement | null

  try {
    const sdk = await getSDK()
    await performRetycTransfer(item, sdk, filesToUpload, {
      recipients,
      appUrl: settings.appUrl,
      expiresDays: settings.expiresDays,
      passphrase,
      onProgress: (p) => {
        if (statusEl) statusEl.textContent = `Uploading "${p.fileName}" (${p.fileIndex + 1}/${p.totalFiles})…`
        if (fillEl) fillEl.style.width = `${Math.round(((p.fileIndex + 1) / p.totalFiles) * 80)}%`
      },
    })
    if (fillEl) fillEl.style.width = '100%'
    if (statusEl) statusEl.textContent = 'Finalizing…'
    clearSelectedFiles()
    showComposeView('done')
    console.info('[Retyc] transfer ready')
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err)
    console.error('[Retyc] transfer failed:', err)

    if (/\b409\b/.test(message)) {
      showComposeView('prepare')
      if (sendButton) sendButton.disabled = false
      showComposePassphraseError('A passphrase is required for recipients without a Retyc account.')
      passphraseInput?.focus()
      return
    }

    if (isAuthError(message)) {
      try { await (await getSDK()).auth.logout() } catch { /* best effort */ }
      _cachedUserInfo = null
      showAuth('auth-unauthenticated')
      showComposeError('Your Retyc session has expired. Please log in again above and retry.')
      return
    }

    showComposeError(message)
  }
}

function backToPrepare(): void {
  clearComposePassphraseError()
  showComposeView('prepare')
  renderSelectedFiles()
  updateSendButtonState()
}

// --- Dropzone wiring ---

function setupDropzone(): void {
  const zone = document.getElementById('dropzone')
  const input = document.getElementById('file-input') as HTMLInputElement | null
  if (!zone || !input) return

  const openPicker = (): void => input.click()
  zone.addEventListener('click', openPicker)
  zone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault()
      openPicker()
    }
  })

  input.addEventListener('change', () => {
    if (input.files) addFiles(input.files)
    input.value = '' // allow re-selecting the same file
  })

  ;(['dragenter', 'dragover'] as const).forEach((evt) => {
    zone.addEventListener(evt, (e) => {
      e.preventDefault()
      e.stopPropagation()
      zone.classList.add('is-dragover')
    })
  })
  ;(['dragleave', 'dragend', 'drop'] as const).forEach((evt) => {
    zone.addEventListener(evt, (e) => {
      e.preventDefault()
      e.stopPropagation()
      zone.classList.remove('is-dragover')
    })
  })
  zone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer
    if (dt?.files) addFiles(dt.files)
  })
}

void Office.onReady(() => {
  document.getElementById('btn-login')?.addEventListener('click', () => { void startLogin() })
  document.getElementById('btn-cancel-login')?.addEventListener('click', cancelLogin)
  document.getElementById('btn-logout')?.addEventListener('click', () => { void logout() })
  document.getElementById('btn-refresh')?.addEventListener('click', () => { void refreshToken() })
  document.getElementById('settings-form')?.addEventListener('submit', (e) => { void saveSettings(e) })

  document.getElementById('btn-send-via-retyc')?.addEventListener('click', () => { void sendViaRetyc() })
  document.getElementById('btn-compose-retry')?.addEventListener('click', backToPrepare)
  document.getElementById('btn-compose-new')?.addEventListener('click', () => {
    backToPrepare()
    void readAuthSnapshot().then((s) => refreshComposeSection(s.authenticated))
  })
  document.getElementById('btn-clear-files')?.addEventListener('click', clearSelectedFiles)

  const recipientsInput = document.getElementById('recipients-input') as HTMLTextAreaElement | null
  recipientsInput?.addEventListener('input', () => {
    _recipientsTouched = true
    renderRecipientsState()
  })

  document.getElementById('btn-recipients-from-outlook')?.addEventListener('click', () => {
    const item = getComposeItem()
    if (item) void fillRecipientsFromOutlook(item, true)
  })

  document.getElementById('btn-toggle-pass')?.addEventListener('click', () => {
    const input = document.getElementById('compose-passphrase') as HTMLInputElement
    input.type = input.type === 'password' ? 'text' : 'password'
  })

  // Re-evaluate the send button state whenever the passphrase changes — without recipients,
  // a long-enough passphrase is what unlocks the upload.
  document.getElementById('compose-passphrase')?.addEventListener('input', updateSendButtonState)

  // Wire collapsible section headers (chevron toggles).
  document.querySelectorAll<HTMLElement>('.collapsible > .section-header').forEach((header) => {
    header.addEventListener('click', () => {
      header.parentElement?.classList.toggle('collapsed')
    })
  })

  setupDropzone()
  renderSelectedFiles()
  renderRecipientsState()

  window.addEventListener('beforeunload', stopPolling)

  void loadStatus()
})
