import { contextBridge, ipcRenderer } from 'electron'
import type {
  AppSettings,
  MasterEntry,
  MasterListState,
  AdoWorkItemDto,
  AdoProjectDto,
  AdoCommentDto,
  OutlookSnapshot,
  OutlookEventDto,
  LocalNote,
  NowPin
} from '../../shared/types'

export interface SyncResult {
  items: AdoWorkItemDto[]
  masterList: MasterListState
  added: number
  updated: number
  removed: number
  errors: string[]
  syncedAs: string[]
}

const api = {
  getSettings: (): Promise<AppSettings> => ipcRenderer.invoke('settings:get'),
  saveSettings: (settings: AppSettings): Promise<AppSettings> =>
    ipcRenderer.invoke('settings:save', settings),

  getMasterList: (): Promise<MasterListState> => ipcRenderer.invoke('master:get'),
  saveMasterList: (state: MasterListState): Promise<MasterListState> =>
    ipcRenderer.invoke('master:save', state),
  updateMasterEntry: (entry: MasterEntry): Promise<MasterListState> =>
    ipcRenderer.invoke('master:updateEntry', entry),
  getLocalNotes: (): Promise<LocalNote[]> => ipcRenderer.invoke('notes:get'),
  saveLocalNotes: (notes: LocalNote[]): Promise<LocalNote[]> =>
    ipcRenderer.invoke('notes:save', notes),
  getNowPins: (): Promise<NowPin[]> => ipcRenderer.invoke('now:get'),
  saveNowPins: (pins: NowPin[]): Promise<NowPin[]> => ipcRenderer.invoke('now:save', pins),

  listProjects: (connectionId: string): Promise<AdoProjectDto[]> =>
    ipcRenderer.invoke('ado:listProjects', connectionId),
  syncMyWorkItems: (): Promise<SyncResult> =>
    ipcRenderer.invoke('ado:syncMyWorkItems'),
  getWorkItem: (payload: {
    connectionId: string
    id: number
  }): Promise<AdoWorkItemDto> => ipcRenderer.invoke('ado:getWorkItem', payload),
  getWorkItemComments: (payload: {
    connectionId: string
    id: number
    project: string
  }): Promise<AdoCommentDto[]> =>
    ipcRenderer.invoke('ado:getWorkItemComments', payload),
  updateWorkItem: (payload: {
    connectionId: string
    id: number
    title?: string
    state?: string
    description?: string
  }): Promise<AdoWorkItemDto> => ipcRenderer.invoke('ado:updateWorkItem', payload),

  openExternal: (url: string): Promise<void> =>
    ipcRenderer.invoke('shell:openExternal', url),
  playAlertSound: (soundName?: string): Promise<void> =>
    ipcRenderer.invoke('sound:play', soundName),
  getAlertsMuted: (): Promise<boolean> => ipcRenderer.invoke('alerts:getMuted'),
  setAlertsMuted: (muted: boolean): Promise<boolean> =>
    ipcRenderer.invoke('alerts:setMuted', muted),
  onAlertsMuted: (callback: (muted: boolean) => void): (() => void) => {
    const listener = (_event: unknown, muted: boolean) => callback(muted)
    ipcRenderer.on('alerts:muted', listener)
    return () => {
      ipcRenderer.removeListener('alerts:muted', listener)
    }
  },

  getOutlook: (): Promise<OutlookSnapshot> => ipcRenderer.invoke('outlook:get'),
  getOutlookDay: (ymd: string): Promise<OutlookEventDto[]> =>
    ipcRenderer.invoke('outlook:day', ymd),
  connectOutlook: (clientId: string): Promise<OutlookSnapshot> =>
    ipcRenderer.invoke('outlook:connect', clientId),
  disconnectOutlook: (): Promise<OutlookSnapshot> =>
    ipcRenderer.invoke('outlook:disconnect'),
  onOutlook: (callback: (snapshot: OutlookSnapshot) => void): (() => void) => {
    const listener = (_event: unknown, snapshot: OutlookSnapshot) => callback(snapshot)
    ipcRenderer.on('outlook:snapshot', listener)
    return () => {
      ipcRenderer.removeListener('outlook:snapshot', listener)
    }
  },
  onMsWebOpenTab: (callback: (id: string) => void): (() => void) => {
    const listener = (_event: unknown, id: string) => callback(id)
    ipcRenderer.on('ms-web:open-tab', listener)
    return () => {
      ipcRenderer.removeListener('ms-web:open-tab', listener)
    }
  },
  getMsWebCounts: (): Promise<{
    outlookUnread: number
    teamsUnread: number
    extra: Record<string, number>
  }> => ipcRenderer.invoke('msWeb:counts'),
  syncMsWebCounts: (): Promise<{
    outlookUnread: number
    teamsUnread: number
    extra: Record<string, number>
  }> => ipcRenderer.invoke('msWeb:syncCounts'),
  onMsWebCounts: (
    callback: (counts: {
      outlookUnread: number
      teamsUnread: number
      extra: Record<string, number>
    }) => void
  ): (() => void) => {
    const listener = (
      _event: unknown,
      counts: {
        outlookUnread: number
        teamsUnread: number
        extra: Record<string, number>
      }
    ) => callback(counts)
    ipcRenderer.on('ms-web:counts', listener)
    return () => {
      ipcRenderer.removeListener('ms-web:counts', listener)
    }
  },
  showMsWeb: (payload: {
    id: string
    url?: string
    bounds: { x: number; y: number; width: number; height: number }
  }): Promise<void> => ipcRenderer.invoke('msWeb:show', payload),
  hideMsWeb: (): Promise<void> => ipcRenderer.invoke('msWeb:hide'),
  focusMsWeb: (): Promise<void> => ipcRenderer.invoke('msWeb:focus'),
  reloadMsWeb: (id: string, url?: string): Promise<void> =>
    ipcRenderer.invoke('msWeb:reload', { id, url })
}

contextBridge.exposeInMainWorld('adoApi', api)

export type AdoApi = typeof api
