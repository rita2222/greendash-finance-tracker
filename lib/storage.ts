import { storageGet, storageSet } from './supabase'

export const windowStorage = {
  async get(key: string) {
    try {
      return await storageGet(key)
    } catch {
      return null
    }
  },
  async set(key: string, value: string) {
    try {
      return await storageSet(key, value)
    } catch {
      return false
    }
  }
}
