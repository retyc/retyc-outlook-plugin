import { RetycSDK } from '@retyc/sdk'
import { OfficeRoamingTokenStore } from './token-store'
import { getApiUrl } from './settings'

let _sdk: RetycSDK | null = null
let _currentApiUrl: string | null = null
const _store = new OfficeRoamingTokenStore()

export async function getSDK(): Promise<RetycSDK> {
  const apiUrl = getApiUrl()
  if (_sdk && _currentApiUrl === apiUrl) return _sdk

  const sdk = new RetycSDK({ apiUrl, tokenStore: _store })
  // Pre-load OIDC config so refresh works on first call after a runtime cold-start.
  await sdk.preload()
  _sdk = sdk
  _currentApiUrl = apiUrl
  return _sdk
}

export function invalidateSDK(): void {
  _sdk = null
  _currentApiUrl = null
}
