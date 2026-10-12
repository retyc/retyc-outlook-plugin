import { isTokenSet, type TokenSet, type TokenStore } from '@retyc/sdk'
import { STORAGE_KEY_TOKENS } from './constants'
import { getStore } from './storage'

// Persists OIDC tokens through the unified store: Office.context.roamingSettings inside
// Outlook (mailbox-scoped, syncs across devices), localStorage when running standalone.
export class OfficeRoamingTokenStore implements TokenStore {
  async get(): Promise<TokenSet | null> {
    const raw = getStore().get(STORAGE_KEY_TOKENS)
    if (typeof raw !== 'string' || !raw) return null
    try {
      const parsed: unknown = JSON.parse(raw)
      if (!isTokenSet(parsed)) {
        await this.clear()
        return null
      }
      return parsed
    } catch {
      await this.clear()
      return null
    }
  }

  async set(tokens: TokenSet): Promise<void> {
    const store = getStore()
    store.set(STORAGE_KEY_TOKENS, JSON.stringify(tokens))
    await store.save()
  }

  async clear(): Promise<void> {
    const store = getStore()
    store.remove(STORAGE_KEY_TOKENS)
    await store.save()
  }
}
