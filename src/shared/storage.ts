// Persistent key/value store backing both settings and tokens.
//
// Inside Outlook: backed by Office.context.roamingSettings (mailbox-scoped, syncs across devices).
// Outside Outlook (browser standalone testing): falls back to window.localStorage so the UI loads
// and persists during the dev session.

interface PersistentStore {
  get(key: string): unknown
  set(key: string, value: unknown): void
  remove(key: string): void
  save(): Promise<void>
}

const LOCAL_STORAGE_PREFIX = 'retyc:'

class RoamingSettingsStore implements PersistentStore {
  constructor(private readonly rs: Office.RoamingSettings) {}

  get(key: string): unknown {
    return this.rs.get(key)
  }

  set(key: string, value: unknown): void {
    this.rs.set(key, value)
  }

  remove(key: string): void {
    this.rs.remove(key)
  }

  save(): Promise<void> {
    return new Promise((resolve, reject) => {
      this.rs.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve()
        else reject(new Error(result.error?.message ?? 'Failed to save roaming settings.'))
      })
    })
  }
}

class LocalStorageStore implements PersistentStore {
  get(key: string): unknown {
    const raw = window.localStorage.getItem(LOCAL_STORAGE_PREFIX + key)
    if (raw === null) return undefined
    try {
      return JSON.parse(raw) as unknown
    } catch {
      return raw
    }
  }

  set(key: string, value: unknown): void {
    window.localStorage.setItem(LOCAL_STORAGE_PREFIX + key, JSON.stringify(value))
  }

  remove(key: string): void {
    window.localStorage.removeItem(LOCAL_STORAGE_PREFIX + key)
  }

  save(): Promise<void> {
    return Promise.resolve()
  }
}

let _store: PersistentStore | null = null
let _warned = false

export function getStore(): PersistentStore {
  if (_store) return _store

  const rs = (typeof Office !== 'undefined' && Office.context)
    ? Office.context.roamingSettings
    : undefined

  if (rs) {
    _store = new RoamingSettingsStore(rs)
  } else {
    if (!_warned) {
      console.warn(
        '[Retyc] Office.context.roamingSettings unavailable — falling back to localStorage. ' +
          'This is expected when opening the page directly in a browser; settings will not roam.',
      )
      _warned = true
    }
    _store = new LocalStorageStore()
  }
  return _store
}
