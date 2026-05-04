import { ref, computed, onMounted, onUnmounted, watch } from 'vue'
import { getSDK } from '../shared/sdk-factory'
import { performRetycTransfer, readInitialRecipients, isAuthError } from '../shared/upload'
import type { UploadProgress } from '../shared/upload'
import { t, getLocale } from '../i18n'

export type ComposeView = 'prepare' | 'uploading' | 'done' | 'error'

export interface UploadStatus {
  phase:         'reading' | 'uploading'
  fileName:      string
  fileIndex:     number  // 1-based for display
  totalFiles:    number
  uploadedBytes: number
  totalBytes:    number
  ratio:         number  // 0..1
}

export interface ExpiryOption {
  label: string
  value: number
}

const EXPIRY_VALUES = [
  { key: 'hour1',   value: 3600 },
  { key: 'hours12', value: 43200 },
  { key: 'day1',    value: 86400 },
  { key: 'days3',   value: 259200 },
  { key: 'days7',   value: 604800 },
  { key: 'days30',  value: 2592000 },
  { key: 'days90',  value: 7776000 },
  { key: 'year1',   value: 31536000 },
] as const

function buildExpiryOptions(): ExpiryOption[] {
  return EXPIRY_VALUES.map(({ key, value }) => ({
    label: t(`transfer.expiryOptions.${key}`),
    value,
  }))
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/

export function isValidEmail(email: string): boolean {
  return EMAIL_RE.test(email)
}

function getComposeItem(): Office.MessageCompose | null {
  if (typeof Office === 'undefined' || !Office.context?.mailbox) return null
  const item = Office.context.mailbox.item as Office.MessageCompose | undefined
  if (item && 'to' in item && item.to && typeof item.to.getAsync === 'function') return item
  return null
}

export interface UseTransferOptions {
  isAuthenticated:  () => boolean
  onAuthError:      () => void
  currentUserEmail: () => string
}

export function useTransfer(options: UseTransferOptions) {
  const composeView = ref<ComposeView>('prepare')
  const selectedFiles = ref<File[]>([])
  const recipients = ref<string[]>([])
  const passphrase = ref('')
  const passphraseError    = ref('')
  const passphraseRequired = ref(false)
  const expirySeconds = ref(604800)
  const expiryOptions = ref<ExpiryOption[]>([{ label: t('transfer.expiryOptions.days7'), value: 604800 }])
  const maxShareSize  = ref<number | null>(null)
  const uploadStatus  = ref<UploadStatus | null>(null)
  const transferError = ref('')

  const recipientsTouched = ref(false)
  let itemChangedHandler: (() => void) | null = null
  let recipientsHandler: (() => void) | null = null
  let recipientsHandlerItem: Office.MessageCompose | null = null

  const isComposeAvailable = computed(() => getComposeItem() !== null && options.isAuthenticated())

  const invalidRecipients = computed(() =>
    recipients.value.filter(email => email.length > 0 && !isValidEmail(email)),
  )
  const validRecipients = computed(() =>
    recipients.value.filter(email => isValidEmail(email)),
  )
  const hasInvalidRecipients = computed(() => invalidRecipients.value.length > 0)

  const canSend = computed(() => {
    const hasFiles = selectedFiles.value.length > 0
    const hasRecip = validRecipients.value.length > 0
    const hasPass = passphrase.value.trim().length >= 8
    return hasFiles && !hasInvalidRecipients.value && (hasRecip || hasPass)
  })

  async function loadCapabilities() {
    const all = buildExpiryOptions()
    try {
      const sdk = await getSDK()
      const caps = await sdk.user.getUploadCapabilities()
      const max = caps.max_share_expiration_time
      expiryOptions.value = max == null
        ? all
        : all.filter(o => o.value <= max)
      maxShareSize.value = caps.max_share_size
    } catch {
      expiryOptions.value = all.filter(o => o.value <= 604800)
    }
    if (!expiryOptions.value.find(o => o.value === expirySeconds.value)) {
      expirySeconds.value = expiryOptions.value[expiryOptions.value.length - 1]?.value ?? 604800
    }
  }

  function isSelf(email: string): boolean {
    const self = options.currentUserEmail().toLowerCase()
    return self.length > 0 && email.toLowerCase() === self
  }

  async function prefillRecipients(force = false) {
    if (!force && recipientsTouched.value) return
    const item = getComposeItem()
    if (!item) return
    try {
      const emails = await readInitialRecipients(item)
      const filtered = emails.map(e => e.toLowerCase()).filter(e => !isSelf(e))
      if (filtered.length > 0) {
        recipients.value = filtered
      }
    } catch { /* silent */ }
  }

  async function syncFromOutlook() {
    await prefillRecipients(true)
  }

  async function refreshComposeIfNeeded() {
    if (!options.isAuthenticated() || !getComposeItem()) {
      composeView.value = 'prepare'
      return
    }
    if (!recipientsTouched.value) await prefillRecipients()
    void loadCapabilities()
    attachRecipientsHandler()
  }

  function addFiles(files: FileList | File[]) {
    for (const f of Array.from(files)) {
      const dup = selectedFiles.value.some(
        s => s.name === f.name && s.size === f.size && s.lastModified === f.lastModified,
      )
      if (!dup) selectedFiles.value.push(f)
    }
  }

  function removeFile(index: number) {
    selectedFiles.value.splice(index, 1)
  }

  function clearFiles() {
    selectedFiles.value = []
  }

  function markRecipientsTouched() {
    recipientsTouched.value = true
  }

  function backToPrepare() {
    passphraseError.value    = ''
    passphraseRequired.value = false
    composeView.value        = 'prepare'
  }

  function resetTransferState() {
    recipientsTouched.value  = false
    selectedFiles.value      = []
    passphrase.value         = ''
    passphraseError.value    = ''
    passphraseRequired.value = false
    composeView.value        = 'prepare'
  }

  async function sendFiles() {
    const item = getComposeItem()
    if (!item) {
      transferError.value = t('transfer.errors.noActiveCompose')
      composeView.value = 'error'
      return
    }
    if (!selectedFiles.value.length) {
      passphraseError.value = t('transfer.errors.dropAtLeastOneFile')
      return
    }
    if (hasInvalidRecipients.value) {
      passphraseError.value = t('transfer.errors.fixInvalidAddresses')
      return
    }

    const pass = passphrase.value.trim() || undefined
    if (pass && pass.length < 8) {
      passphraseError.value = t('transfer.errors.passphraseTooShort')
      return
    }
    if (!validRecipients.value.length && !pass) {
      passphraseError.value = t('transfer.errors.addRecipientOrPass')
      return
    }

    passphraseError.value = ''
    composeView.value = 'uploading'
    uploadStatus.value = null
    const filesToUpload = [...selectedFiles.value]
    passphrase.value = ''

    try {
      const sdk = await getSDK()
      await performRetycTransfer(item, sdk, filesToUpload, {
        recipients: validRecipients.value,
        expires: expirySeconds.value,
        passphrase: pass,
        onProgress: (p: UploadProgress) => {
          uploadStatus.value = {
            phase:         p.phase,
            fileName:      p.fileName,
            fileIndex:     p.fileIndex + 1,
            totalFiles:    p.totalFiles,
            uploadedBytes: p.uploadedBytes,
            totalBytes:    p.totalBytes,
            ratio:         p.ratio,
          }
        },
      })
      clearFiles()
      composeView.value = 'done'
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      if (/\b409\b/.test(msg)) {
        composeView.value    = 'prepare'
        passphraseRequired.value = true
        passphraseError.value    = t('transfer.errors.passphraseRequired409')
        return
      }
      if (isAuthError(msg)) {
        try {
          await (await getSDK()).auth.logout()
        } catch { /* best effort */ }
        options.onAuthError()
        return
      }
      transferError.value = msg
      composeView.value = 'error'
    }
  }

  function attachRecipientsHandler() {
    const item = getComposeItem()
    if (!item || recipientsHandler) return
    recipientsHandler = () => {
      const current = getComposeItem()
      if (!current) return
      void readInitialRecipients(current).then(emails => {
        recipients.value = emails
          .map(e => e.trim().toLowerCase())
          .filter(e => e.length > 0 && !isSelf(e))
      })
    }
    try {
      item.addHandlerAsync(Office.EventType.RecipientsChanged, recipientsHandler)
      recipientsHandlerItem = item
    } catch { /* not supported */ }
  }

  function setupOfficeEventHandlers() {
    if (typeof Office === 'undefined' || !Office.context?.mailbox) return

    if (!itemChangedHandler) {
      itemChangedHandler = () => {
        if (recipientsHandler && recipientsHandlerItem) {
          try {
            recipientsHandlerItem.removeHandlerAsync(Office.EventType.RecipientsChanged)
          } catch { /* best effort — the previous item may already be gone */ }
        }
        recipientsHandler = null
        recipientsHandlerItem = null
        resetTransferState()
        void refreshComposeIfNeeded()
      }
      try {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChangedHandler)
      } catch { /* not supported */ }
    }

    attachRecipientsHandler()
  }

  function teardownOfficeEventHandlers() {
    if (typeof Office === 'undefined' || !Office.context?.mailbox) return

    if (itemChangedHandler) {
      try {
        Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged)
      } catch { /* best effort */ }
      itemChangedHandler = null
    }
    if (recipientsHandler && recipientsHandlerItem) {
      try {
        recipientsHandlerItem.removeHandlerAsync(Office.EventType.RecipientsChanged)
      } catch { /* best effort */ }
      recipientsHandler = null
      recipientsHandlerItem = null
    }
  }

  // Refresh expiry option labels when locale changes.
  watch(() => getLocale(), () => {
    if (expiryOptions.value.length === 0) return
    const allowedValues = new Set(expiryOptions.value.map(o => o.value))
    expiryOptions.value = buildExpiryOptions().filter(o => allowedValues.has(o.value))
  })

  onMounted(() => {
    setupOfficeEventHandlers()
  })

  onUnmounted(() => {
    teardownOfficeEventHandlers()
  })

  return {
    composeView,
    selectedFiles,
    recipients,
    passphrase,
    passphraseError,
    passphraseRequired,
    expirySeconds,
    expiryOptions,
    maxShareSize,
    uploadStatus,
    transferError,
    isComposeAvailable,
    invalidRecipients,
    canSend,
    addFiles,
    removeFile,
    clearFiles,
    markRecipientsTouched,
    syncFromOutlook,
    refreshComposeIfNeeded,
    sendFiles,
    backToPrepare,
  }
}
