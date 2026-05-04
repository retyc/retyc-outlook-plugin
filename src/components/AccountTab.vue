<script setup lang="ts">
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { injectAuth } from '../composables/useAuth'
import { getSDK } from '../shared/sdk-factory'
import { setLocale, type AppLocale } from '../i18n'

const auth = injectAuth()
const { t, locale } = useI18n()

const now = ref(Math.floor(Date.now() / 1000))
let ticker: ReturnType<typeof setInterval> | null = null

function startTicker() {
  if (ticker !== null) return
  ticker = setInterval(() => { now.value = Math.floor(Date.now() / 1000) }, 1000)
}

onUnmounted(() => {
  if (ticker !== null) { clearInterval(ticker); ticker = null }
})

const expiryText = computed(() => {
  const tk = auth.tokenExpiry.value
  if (!tk) return ''
  const s = tk - now.value
  if (s <= 0)  return t('account.expiry.expired')
  if (s < 3600) startTicker()
  if (s < 60)  return t('account.expiry.seconds', { value: s })
  const m = Math.floor(s / 60)
  if (m < 60)  return t('account.expiry.minutes', { value: m })
  const h = Math.floor(m / 60)
  const rm = m % 60
  if (h < 24)  return rm > 0
    ? t('account.expiry.hoursMins', { hours: h, minutes: rm })
    : t('account.expiry.hours', { value: h })
  return t('account.expiry.days', { value: Math.floor(h / 24) })
})

interface Quota {
  usedStorage:    number
  maxStorage:     number
  countShare:     number
  maxCountShare:  number | null
}

const quota = ref<Quota | null>(null)

function formatSize(bytes: number): string {
  if (bytes < 1024)        return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  if (bytes < 1024 ** 3)   return `${(bytes / 1024 / 1024).toFixed(1)} MB`
  return `${(bytes / 1024 ** 3).toFixed(2)} GB`
}

const storagePercent = computed(() => {
  if (!quota.value || quota.value.maxStorage === 0) return 0
  return Math.min(100, Math.round((quota.value.usedStorage / quota.value.maxStorage) * 100))
})

const transferPercent = computed(() => {
  if (!quota.value || quota.value.maxCountShare === null) return null
  return Math.min(100, Math.round((quota.value.countShare / quota.value.maxCountShare) * 100))
})

const localeItems = computed(() => [
  { label: t('account.languageEnglish'), value: 'en' as AppLocale },
  { label: t('account.languageFrench'),  value: 'fr' as AppLocale },
])

const selectedLocale = computed<AppLocale>({
  get: () => (locale.value === 'fr' ? 'fr' : 'en'),
  set: (next) => setLocale(next),
})

onMounted(async () => {
  try {
    const sdk = await getSDK()
    const q = await sdk.user.getUserQuota()
    quota.value = {
      usedStorage:   q.used_storage,
      maxStorage:    q.max_storage,
      countShare:    q.count_share,
      maxCountShare: q.max_count_share,
    }
  } catch { /* silent — quota is informational */ }
})
</script>

<template>
  <div class="flex flex-col gap-4 px-3 py-4">
    <!-- User info -->
    <UCard>
      <div class="flex items-center justify-between gap-3">
        <div class="flex flex-col min-w-0">
          <span v-if="auth.userFullName.value" class="text-sm font-semibold text-slate-900 truncate">
            {{ auth.userFullName.value }}
          </span>
          <span class="text-xs text-neutral-500 truncate">{{ auth.userEmail.value }}</span>
          <span v-if="expiryText" class="text-[11px] text-neutral-400 mt-0.5">{{ expiryText }}</span>
        </div>
        <UButton
          color="neutral"
          variant="ghost"
          icon="i-lucide-refresh-ccw"
          size="sm"
          :title="t('account.refreshTokenTitle')"
          @click="auth.refreshToken"
        />
      </div>
    </UCard>

    <!-- Quota -->
    <UCard v-if="quota">
      <div class="flex flex-col gap-3">
        <!-- Storage -->
        <div class="flex flex-col gap-1">
          <div class="flex justify-between text-xs text-neutral-500">
            <span>{{ t('account.storage') }}</span>
            <span>{{ formatSize(quota.usedStorage) }} / {{ formatSize(quota.maxStorage) }}</span>
          </div>
          <UProgress :model-value="storagePercent" :color="storagePercent >= 90 ? 'error' : 'primary'" size="sm"/>
        </div>

        <!-- Transfers (only if capped) -->
        <div v-if="quota.maxCountShare !== null" class="flex flex-col gap-1">
          <div class="flex justify-between text-xs text-neutral-500">
            <span>{{ t('account.transfers') }}</span>
            <span>{{ quota.countShare }} / {{ quota.maxCountShare }}</span>
          </div>
          <UProgress :model-value="transferPercent ?? 0" :color="(transferPercent ?? 0) >= 90 ? 'error' : 'primary'" size="sm"/>
        </div>
      </div>
    </UCard>

    <!-- Language -->
    <UFormField :label="t('account.language')">
      <USelect
        v-model="selectedLocale"
        :items="localeItems"
        class="w-full"
      />
    </UFormField>

    <UButton
      color="neutral"
      size="md"
      variant="subtle"
      block
      icon="i-lucide-log-out"
      @click="auth.logout"
    >
      {{ t('account.logOut') }}
    </UButton>
  </div>
</template>
