import { app, safeStorage } from 'electron'

const PREFIX = 'safeStorage:v1:'

export function isEncryptedSecret(value: string): boolean {
  return String(value ?? '').startsWith(PREFIX)
}

export function encryptSecret(plain: string): string {
  const value = String(plain ?? '')
  if (!value || isEncryptedSecret(value)) return value
  if (!app.isReady() || !safeStorage.isEncryptionAvailable()) return value
  return PREFIX + safeStorage.encryptString(value).toString('base64')
}

export function decryptSecret(value: string): string {
  const raw = String(value ?? '')
  if (!raw || !isEncryptedSecret(raw)) return raw
  if (!safeStorage.isEncryptionAvailable()) {
    console.error('safeStorage unavailable; cannot decrypt secret')
    return ''
  }
  try {
    return safeStorage.decryptString(Buffer.from(raw.slice(PREFIX.length), 'base64'))
  } catch (err) {
    console.error('Failed to decrypt secret', err)
    return ''
  }
}

export function encryptionAvailable(): boolean {
  return app.isReady() && safeStorage.isEncryptionAvailable()
}
