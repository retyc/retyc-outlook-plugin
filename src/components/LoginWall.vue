<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { injectAuth } from '../composables/useAuth'

const auth = injectAuth()
const { t } = useI18n()

async function onLogin() {
  await auth.startLogin()
}
</script>

<template>
  <div class="flex flex-col gap-5 px-5 py-6">
    <div class="flex flex-col items-center gap-4 text-center">
      <img src="/assets/icon-80.png" alt="Retyc" class="w-12 h-12"/>
      <div>
        <h1 class="text-3xl font-bold text-cornflower-blue-900">{{ t('app.name') }}</h1>
        <p class="text-sm text-neutral-400 mt-0.5">{{ t('app.tagline') }}</p>
      </div>
    </div>

    <div v-if="auth.authState.value === 'loading'" class="flex justify-center py-8">
      <img src="/assets/icon-80.png" alt="" class="spinner-icon w-12 h-12"/>
    </div>

    <div v-else-if="auth.authState.value === 'unauthenticated'" class="flex flex-col items-center gap-4">
      <UButton
        id="btn-login"
        size="xl"
        color="primary"
        block
        icon="i-lucide-log-in"
        @click="onLogin"
      >
        {{ t('login.button') }}
      </UButton>
      <USeparator/>
      <p class="text-sm text-neutral-500 text-center">
        {{ t('login.newToRetyc') }}
        <UButton to="https://retyc.com/" external target="_blank" variant="link" color="secondary" size="sm">{{ t('login.signUp') }}</UButton>
      </p>
    </div>

    <div v-else-if="auth.authState.value === 'device-flow'" class="flex flex-col items-center gap-4">
      <USeparator/>
      <p class="text-sm text-neutral-500">{{ t('login.deviceFlowIntro') }}</p>
      <UButton
        :to="auth.deviceFlowUrl.value"
        external
        target="_blank"
        size="xl"
        block
        trailing-icon="i-lucide-external-link"
      >
        {{ t('login.openSignInPage') }}
      </UButton>
      <p class="text-sm text-neutral-500">{{ t('login.enterCodePrompt') }}</p>
      <div class="bg-neutral-100 rounded-lg px-6 py-3">
        <span class="font-mono text-2xl font-bold text-cornflower-blue-800 tracking-widest">
          {{ auth.userCode.value }}
        </span>
      </div>
      <p class="text-sm text-neutral-500 flex items-center gap-2">
        {{ t('login.waiting') }}
        <UIcon name="i-lucide-loader-circle" class="animate-spin"/>
      </p>
      <UButton color="neutral" variant="subtle" size="lg" @click="auth.cancelLogin">{{ t('login.cancel') }}</UButton>
    </div>
  </div>
</template>
