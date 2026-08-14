import type { AdoApi } from '../preload/index'

declare global {
  interface Window {
    adoApi: AdoApi
  }
}

export {}
