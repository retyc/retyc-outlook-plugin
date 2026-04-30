declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent<object, object, unknown>
  export default component
}

declare module 'virtual:lucide-subset' {
  import type { IconifyJSON } from '@iconify/types'
  const collection: IconifyJSON
  export default collection
}

interface ImportMetaEnv {
  readonly VITE_RETYC_API_URL?: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}
