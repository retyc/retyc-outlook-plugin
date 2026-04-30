import { Buffer } from 'buffer'
import { addCollection } from '@iconify/vue'
import lucide from 'virtual:lucide-subset'
import { createApp } from 'vue'
import ui from '@nuxt/ui/vue-plugin'
import App from './App.vue'
import { i18n, setLocale, detectInitialLocale } from './i18n'
import './assets/custom.css'

addCollection(lucide)

if (typeof globalThis.Buffer === 'undefined') {
  (globalThis as typeof globalThis & { Buffer: typeof Buffer }).Buffer = Buffer
}

void Office.onReady(() => {
  setLocale(detectInitialLocale())
  createApp(App).use(i18n).use(ui).mount('#app')
})
