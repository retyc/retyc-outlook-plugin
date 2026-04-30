<script setup lang="ts">
import { onMounted, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import DropZone from './DropZone.vue'
import { injectAuth } from '../composables/useAuth'
import { useTransfer } from '../composables/useTransfer'
import { useDropOverlay } from '../composables/useDropOverlay'
import { t as gt } from '../i18n'

const auth = injectAuth()
const { t } = useI18n()

const transfer = useTransfer({
  isAuthenticated:  () => auth.isAuthenticated.value,
  currentUserEmail: () => auth.userEmail.value,
  onAuthError: () => {
    auth.forceLogout()
    auth.notify(gt('toast.sessionExpired'), 'error')
  },
})

const overlay = useDropOverlay({
  isActive: () => transfer.composeView.value === 'prepare' && transfer.isComposeAvailable.value,
  onFiles: files => transfer.addFiles(files),
})

const showPass        = ref(false)
const usePassphrase   = ref(false)

watch(() => transfer.passphraseRequired.value, required => {
  if (required) usePassphrase.value = true
})

watch(usePassphrase, on => {
  if (!on) {
    transfer.passphrase.value    = ''
    transfer.passphraseError.value = ''
  }
})

function formatSize(bytes: number): string {
  if (bytes < 1024)        return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  if (bytes < 1024 ** 3)   return `${(bytes / 1024 / 1024).toFixed(1)} MB`
  return `${(bytes / 1024 ** 3).toFixed(2)} GB`
}

function handleNewTransfer() {
  transfer.backToPrepare()
  void transfer.refreshComposeIfNeeded()
}

onMounted(async () => {
  await transfer.refreshComposeIfNeeded()
})
</script>

<template>
  <div class="flex flex-col gap-4 px-3 py-3 relative">

    <div v-if="!transfer.isComposeAvailable.value" class="flex flex-col items-center gap-3 py-8 text-center text-neutral-400">
      <UIcon name="i-lucide-mail-plus" class="w-8 h-8"/>
      <p class="text-sm">{{ t('transfer.openComposeHint') }}</p>
    </div>

    <template v-else-if="transfer.composeView.value === 'prepare'">
      <p class="text-xs text-neutral-400">
        {{ t('transfer.encryptedNote') }}
        <i18n-t keypath="transfer.notThroughOutlook" tag="span">
          <template #neverThrough>
            <strong class="text-neutral-600">{{ t('transfer.notThroughOutlookStrong') }}</strong>
          </template>
        </i18n-t>
      </p>

      <UFormField
        :label="t('transfer.filesLabel')"
        :hint="transfer.maxShareSize.value ? t('transfer.maxSize', { size: formatSize(transfer.maxShareSize.value) }) : undefined"
      >
        <DropZone
          :files="transfer.selectedFiles.value"
          @add-files="transfer.addFiles"
          @remove-file="transfer.removeFile"
          @clear-files="transfer.clearFiles"
        />
      </UFormField>

      <UCard v-if="transfer.recipients.value.length > 0" :ui="{body: 'p-2'}">
        <div class="flex flex-wrap gap-1">
          <UBadge v-for="r in transfer.recipients.value" :key="r" size="sm" color="neutral" variant="subtle">
            {{r}}
          </UBadge>
        </div>
      </UCard>

      <UCollapsible class="border border-neutral-200 rounded-lg overflow-hidden">
        <template #default="{ open }">
          <UButton
            color="neutral"
            variant="ghost"
            block
            :trailing-icon="open ? 'i-lucide-chevron-up' : 'i-lucide-chevron-down'"
            :ui="{ base: 'rounded-none justify-between w-full' }"
          >
            {{ t('transfer.options') }}
          </UButton>
        </template>
        <template #content>
          <div class="flex flex-col gap-4 px-3 pb-3 pt-2">
            <div class="flex flex-col gap-2">
              <USwitch
                v-model="usePassphrase"
                :label="t('transfer.usePassphrase')"
                :description="t('transfer.usePassphraseHint')"
              />
              <UFormField
                v-if="usePassphrase"
                :error="transfer.passphraseError.value || undefined"
              >
                <UInput
                  v-model="transfer.passphrase.value"
                  :type="showPass ? 'text' : 'password'"
                  :placeholder="t('transfer.passphrasePlaceholder')"
                  autocomplete="off"
                  class="w-full"
                  :ui="{ trailing: 'pr-1' }"
                >
                  <template #trailing>
                    <UButton
                      color="neutral"
                      variant="ghost"
                      :icon="showPass ? 'i-lucide-eye-off' : 'i-lucide-eye'"
                      size="xs"
                      @click="showPass = !showPass"
                    />
                  </template>
                </UInput>
              </UFormField>
            </div>

            <UFormField :label="t('transfer.expiry')">
              <USelect
                v-model="transfer.expirySeconds.value"
                :items="transfer.expiryOptions.value"
                class="w-full"
              />
            </UFormField>
          </div>
        </template>
      </UCollapsible>

      <UButton
        size="xl"
        color="primary"
        block
        :disabled="!transfer.canSend.value"
        icon="i-lucide-send"
        @click="transfer.sendFiles"
      >
        {{ t('transfer.sendButton') }}
      </UButton>
    </template>

    <div v-else-if="transfer.composeView.value === 'uploading'" class="flex flex-col items-center gap-3 py-4">
      <img src="/assets/icon-80.png" alt="" class="spinner-icon w-12 h-12"/>

      <div v-if="transfer.uploadStatus.value" class="w-full flex flex-col gap-2">
        <!-- Per-file row -->
        <div class="flex justify-between items-baseline gap-2 text-sm text-neutral-700">
          <span class="truncate" :title="transfer.uploadStatus.value.fileName">
            {{ transfer.uploadStatus.value.phase === 'reading'
              ? t('transfer.readingFile',   { name: transfer.uploadStatus.value.fileName })
              : t('transfer.uploadingFile', { name: transfer.uploadStatus.value.fileName }) }}
          </span>
          <span class="text-xs text-neutral-400 shrink-0">
            {{ transfer.uploadStatus.value.fileIndex }} / {{ transfer.uploadStatus.value.totalFiles }}
          </span>
        </div>

        <!-- Total row -->
        <UProgress :model-value="Math.round(transfer.uploadStatus.value.ratio * 100)"/>
        <div class="flex justify-between text-xs text-neutral-500">
          <span>{{ formatSize(transfer.uploadStatus.value.uploadedBytes) }} / {{ formatSize(transfer.uploadStatus.value.totalBytes) }}</span>
          <span>{{ Math.round(transfer.uploadStatus.value.ratio * 100) }}%</span>
        </div>
      </div>

      <p class="text-xs text-neutral-400 text-center">{{ t('transfer.pleaseWait') }}</p>
    </div>

    <div v-else-if="transfer.composeView.value === 'done'" class="flex flex-col items-center gap-3 py-4">
      <UIcon name="i-lucide-check-circle-2" class="w-12 h-12 text-success-500"/>
      <p class="font-semibold text-success-600">{{ t('transfer.transferReady') }}</p>
      <p class="text-sm text-neutral-500 text-center">
        {{ t('transfer.linkAppended') }}<br/>
        <i18n-t keypath="transfer.clickSendInOutlook" tag="span">
          <template #sendBold><strong>{{ t('transfer.clickSendInOutlookBold') }}</strong></template>
        </i18n-t>
      </p>
      <UButton color="neutral" size="xl" block @click="handleNewTransfer">{{ t('transfer.prepareAnother') }}</UButton>
    </div>

    <div v-else-if="transfer.composeView.value === 'error'" class="flex flex-col items-center gap-3 py-4">
      <UIcon name="i-lucide-x-circle" class="w-12 h-12 text-error-500"/>
      <p class="font-semibold text-error-600">{{ t('transfer.uploadFailed') }}</p>
      <p class="text-sm text-neutral-500 text-center">{{ transfer.transferError.value }}</p>
      <UButton color="neutral" size="xl" block @click="transfer.backToPrepare">{{ t('transfer.back') }}</UButton>
    </div>

    <Transition name="fade">
      <div
        v-if="overlay.showOverlay.value"
        class="drop-overlay"
        @dragover.prevent
        @drop="overlay.onOverlayDrop"
      >
        <UIcon name="i-lucide-upload" class="w-10 h-10"/>
        <p>{{ t('transfer.dropFilesHere') }}</p>
      </div>
    </Transition>
  </div>
</template>

<style scoped>
.fade-enter-active, .fade-leave-active { transition: opacity 0.1s; }
.fade-enter-from, .fade-leave-to       { opacity: 0; }

.drop-overlay {
  position: fixed; inset: 8px; z-index: 500;
  background: color-mix(in srgb, var(--color-primary-50) 94%, transparent);
  border: 2px dashed var(--color-primary-400); border-radius: 12px;
  display: flex; flex-direction: column; align-items: center; justify-content: center;
  gap: 12px; color: var(--color-primary-400); font-size: 15px; font-weight: 600; pointer-events: all;
}
</style>
