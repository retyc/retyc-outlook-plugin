// Typed wrapper around the persistent store (roamingSettings inside Outlook,
// localStorage when running standalone in a browser).

import {
  DEFAULT_API_URL,
  DEFAULT_APP_URL,
  DEFAULT_AUTO_SEND,
  DEFAULT_EXPIRES_DAYS,
  STORAGE_KEY_API_URL,
  STORAGE_KEY_APP_URL,
  STORAGE_KEY_AUTO_SEND,
  STORAGE_KEY_EXPIRES_DAYS,
} from './constants'
import { getStore } from './storage'

export interface RetycSettings {
  apiUrl: string
  appUrl: string
  expiresDays: number
  autoSend: boolean
}

export function saveRoamingSettings(): Promise<void> {
  return getStore().save()
}

export function getApiUrl(): string {
  const raw = getStore().get(STORAGE_KEY_API_URL)
  return typeof raw === 'string' && raw ? raw : DEFAULT_API_URL
}

export function getAppUrl(): string {
  const raw = getStore().get(STORAGE_KEY_APP_URL)
  return typeof raw === 'string' && raw ? raw : DEFAULT_APP_URL
}

export function getExpiresDays(): number {
  const raw = getStore().get(STORAGE_KEY_EXPIRES_DAYS)
  return typeof raw === 'number' && Number.isFinite(raw) ? raw : DEFAULT_EXPIRES_DAYS
}

export function getAutoSend(): boolean {
  const raw = getStore().get(STORAGE_KEY_AUTO_SEND)
  return typeof raw === 'boolean' ? raw : DEFAULT_AUTO_SEND
}

export function getAllSettings(): RetycSettings {
  return {
    apiUrl: getApiUrl(),
    appUrl: getAppUrl(),
    expiresDays: getExpiresDays(),
    autoSend: getAutoSend(),
  }
}

export async function setAllSettings(settings: RetycSettings): Promise<void> {
  const store = getStore()
  store.set(STORAGE_KEY_API_URL, settings.apiUrl)
  store.set(STORAGE_KEY_APP_URL, settings.appUrl)
  store.set(STORAGE_KEY_EXPIRES_DAYS, settings.expiresDays)
  store.set(STORAGE_KEY_AUTO_SEND, settings.autoSend)
  await store.save()
}
