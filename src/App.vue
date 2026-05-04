<script setup lang="ts">
import { computed, defineAsyncComponent, onMounted, provide } from 'vue'
import { useI18n } from 'vue-i18n'
import { en as nuxtEn, fr as nuxtFr } from '@nuxt/ui/locale'
import { useAuth, authInjectionKey } from './composables/useAuth'

const LoginWall = defineAsyncComponent(() => import('./components/LoginWall.vue'))
const TransferTab = defineAsyncComponent(() => import('./components/TransferTab.vue'))
const AccountTab = defineAsyncComponent(() => import('./components/AccountTab.vue'))

const auth = useAuth()
provide(authInjectionKey, auth)

const { t, locale } = useI18n()

const tabItems = computed(() => [
  { label: t('tabs.transfer'), slot: 'transfer' as const },
  { label: t('tabs.account'),  slot: 'account'  as const },
])

const nuxtLocale = computed(() => (locale.value === 'fr' ? nuxtFr : nuxtEn))

onMounted(async () => {
  await auth.loadStatus()
})
</script>

<template>
  <UApp :locale="nuxtLocale">
    <LoginWall v-if="!auth.isAuthenticated.value"/>

    <UTabs
      v-else
      :items="tabItems"
      class="flex flex-col h-full"
      color="neutral"
      variant="link"
      :ui="{ list: 'sticky top-0 z-10 bg-white border-b border-slate-200' }"
    >
      <template #transfer>
        <TransferTab/>
      </template>
      <template #account>
        <AccountTab/>
      </template>
    </UTabs>
  </UApp>
</template>
