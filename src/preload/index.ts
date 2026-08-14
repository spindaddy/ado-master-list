import { contextBridge, ipcRenderer } from 'electron'
import type {
  AppSettings,
  MasterEntry,
  MasterListState,
  AdoWorkItemDto,
  AdoProjectDto,
  AdoCommentDto,
  OutlookSnapshot
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

  getOutlook: (): Promise<OutlookSnapshot> => ipcRenderer.invoke('outlook:get'),
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
  }
}

contextBridge.exposeInMainWorld('adoApi', api)

export type AdoApi = typeof api
