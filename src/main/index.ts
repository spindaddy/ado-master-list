import { app, BrowserWindow, ipcMain, shell, session, Notification, protocol, net } from 'electron'
import { join } from 'path'
import { randomUUID } from 'crypto'
import Store from 'electron-store'
import {
  fetchMyWorkItems,
  fetchWorkItem,
  fetchProjects,
  fetchWorkItemComments,
  updateWorkItem,
  normalizeOrganization,
  isExcludedWorkItemState,
  type AdoCredentials,
  type AdoWorkItem
} from './ado'
import { playNamedAlert } from './alertSounds'
import {
  emptyOutlookSnapshot,
  loadOutlookSnapshot,
  refreshOutlookTokens,
  signInToOutlook,
  type OutlookTokens
} from './outlook'
import type {
  AppSettings,
  AdoOrgConnection,
  AdoProjectDto,
  AdoWorkItemDto,
  MasterEntry,
  MasterListState,
  OutlookSnapshot
} from '../../shared/types'
import { normalizeAlertSound } from '../../shared/types'

type WorkItemWithConn = AdoWorkItem & { connectionId: string }

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'adoimg',
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
      stream: true,
      bypassCSP: true
    }
  }
])

let mainWindow: BrowserWindow | null = null

const store = new Store<{
  settings: AppSettings & Record<string, unknown>
  masterList: MasterListState
  outlookAuth: OutlookTokens | null
  outlookSeen: { mailIds: string[]; meetingKeys: string[]; seeded: boolean }
}>({
  name: 'ado-master-workitems',
  defaults: {
    settings: {
      connections: [],
      autoSyncMinutes: 0,
      alertSound: 'Alarm',
      outlookClientId: '',
      outlookCalendar: true,
      outlookMailAlerts: true,
      outlookMeetingAlertMinutes: 5
    },
    masterList: {
      entries: []
    },
    outlookAuth: null,
    outlookSeen: { mailIds: [], meetingKeys: [], seeded: false }
  }
})

function newConnectionId(): string {
  return randomUUID()
}

function normalizeConnection(raw: Partial<AdoOrgConnection>): AdoOrgConnection | null {
  let organization = (raw.organization ?? '').trim()
  const pat = (raw.pat ?? '').trim()
  if (!organization && !pat) return null
  if (organization) {
    try {
      organization = normalizeOrganization(organization)
    } catch {
      // keep raw
    }
  }
  const enabledProjects = Array.isArray(raw.enabledProjects)
    ? [...new Set(raw.enabledProjects.map((p) => p.trim()).filter(Boolean))]
    : []
  return {
    id: raw.id?.trim() || newConnectionId(),
    organization,
    pat,
    enabledProjects,
    active: raw.active !== false,
    label: raw.label?.trim() || undefined
  }
}

function migrateLegacySettings(raw: Record<string, unknown>): AppSettings {
  const existing = Array.isArray(raw.connections)
    ? (raw.connections as Partial<AdoOrgConnection>[])
        .map(normalizeConnection)
        .filter((c): c is AdoOrgConnection => Boolean(c))
    : []

  // Legacy single-org fields → one connection
  const legacyOrg = String(raw.organization ?? '').trim()
  const legacyPat = String(raw.pat ?? '').trim()
  const legacyProjects = Array.isArray(raw.enabledProjects)
    ? (raw.enabledProjects as string[])
    : []

  let connections = existing
  if (connections.length === 0 && (legacyOrg || legacyPat)) {
    const migrated = normalizeConnection({
      organization: legacyOrg,
      pat: legacyPat,
      enabledProjects: legacyProjects
    })
    if (migrated) connections = [migrated]
  }

  const minutes = Number(raw.autoSyncMinutes)
  const allowed = new Set([0, 5, 15, 30, 60])
  const meetMins = Number(raw.outlookMeetingAlertMinutes)
  const meetAllowed = new Set([0, 5, 10, 15])

  return {
    connections,
    autoSyncMinutes: allowed.has(minutes) ? minutes : 0,
    alertSound: normalizeAlertSound(raw.alertSound),
    outlookClientId: String(raw.outlookClientId ?? '').trim(),
    outlookCalendar: raw.outlookCalendar !== false,
    outlookMailAlerts: raw.outlookMailAlerts !== false,
    outlookMeetingAlertMinutes: meetAllowed.has(meetMins) ? meetMins : 5
  }
}

function normalizeSettings(settings: AppSettings | Record<string, unknown>): AppSettings {
  return migrateLegacySettings(settings as Record<string, unknown>)
}

// One-time migrate persisted settings on load
try {
  store.set('settings', normalizeSettings(store.get('settings') as Record<string, unknown>))
} catch (err) {
  console.error('Failed to migrate settings', err)
}

// Carry over org/PAT from the older projects-based store if present
try {
  const legacy = new Store({ name: 'config' })
  const legacySettings = legacy.get('settings') as Record<string, unknown> | undefined
  const current = normalizeSettings(store.get('settings') as Record<string, unknown>)
  if (
    legacySettings &&
    current.connections.length === 0 &&
    (legacySettings.organization || legacySettings.pat)
  ) {
    store.set(
      'settings',
      normalizeSettings({
        ...current,
        organization: legacySettings.organization,
        pat: legacySettings.pat,
        enabledProjects: legacySettings.enabledProjects
      })
    )
  }
} catch (err) {
  console.error('Failed to import legacy settings', err)
}

function masterEntryId(organization: string, workItemId: number): string {
  return `${organization}:${workItemId}`
}

function toDto(item: WorkItemWithConn): AdoWorkItemDto {
  return {
    id: item.id,
    title: item.title,
    workItemType: item.workItemType,
    state: item.state,
    project: item.project,
    organization: item.organization,
    connectionId: item.connectionId,
    assignedTo: item.assignedTo,
    createdBy: item.createdBy,
    createdDate: item.createdDate,
    changedBy: item.changedBy,
    changedDate: item.changedDate,
    description: item.description,
    webUrl: item.webUrl
  }
}

function previousById(existing: MasterEntry[]): Map<string, MasterEntry> {
  const byId = new Map<string, MasterEntry>()
  for (const e of existing) {
    byId.set(e.id, e)
    if (e.organization) {
      byId.set(masterEntryId(e.organization, e.workItemId), e)
    }
    byId.set(String(e.workItemId), e)
  }
  return byId
}

function entryFromWorkItem(
  item: WorkItemWithConn,
  prev: MasterEntry | undefined,
  now: string
): MasterEntry {
  const id = masterEntryId(item.organization, item.id)
  return {
    id,
    workItemId: item.id,
    title: item.title,
    workItemType: item.workItemType,
    state: item.state,
    project: item.project,
    organization: item.organization,
    connectionId: item.connectionId,
    assignedTo: item.assignedTo,
    createdBy: item.createdBy,
    createdDate: item.createdDate,
    changedBy: item.changedBy,
    description: item.description,
    webUrl: item.webUrl,
    changedDate: item.changedDate,
    notes: prev?.notes ?? '',
    status: prev?.status && prev.status !== 'backlog' ? prev.status : 'active',
    priority: prev?.priority ?? 3,
    addedAt: prev?.addedAt ?? now
  }
}

/** Wipe the stored list and write only this sync's items. */
function replaceMasterFromSync(items: WorkItemWithConn[]): {
  list: MasterListState
  added: number
  updated: number
  removed: number
  newTickets: MasterEntry[]
  hadPriorList: boolean
} {
  const existing = store.get('masterList').entries
  const prevById = previousById(existing)
  const now = new Date().toISOString()

  const next: MasterEntry[] = []
  const newTickets: MasterEntry[] = []
  const seen = new Set<string>()
  let added = 0
  let updated = 0

  for (const item of items) {
    if (isExcludedWorkItemState(item.state)) continue
    const id = masterEntryId(item.organization, item.id)
    if (seen.has(id)) continue
    seen.add(id)
    const prev =
      prevById.get(id) ??
      existing.find(
        (e) => e.workItemId === item.id && e.organization === item.organization
      )
    const entry = entryFromWorkItem(item, prev, now)
    if (prev) updated += 1
    else {
      added += 1
      newTickets.push(entry)
    }
    next.push(entry)
  }

  const list = { entries: next }
  store.set('masterList', list)
  return {
    list,
    added,
    updated,
    removed: Math.max(0, existing.length - next.length),
    newTickets,
    hadPriorList: existing.length > 0
  }
}

function getConnections(): AdoOrgConnection[] {
  return normalizeSettings(store.get('settings') as Record<string, unknown>).connections
}

function getConnection(connectionId: string): AdoOrgConnection {
  const conn = getConnections().find((c) => c.id === connectionId)
  if (!conn) throw new Error('Organization connection not found. Re-save Settings.')
  if (!conn.organization || !conn.pat) {
    throw new Error(`Connection "${conn.label || conn.organization || conn.id}" is missing org or PAT.`)
  }
  return conn
}

function credentialsFor(conn: AdoOrgConnection): AdoCredentials {
  return {
    organization: conn.organization,
    pat: conn.pat
  }
}

async function syncAllConnections(): Promise<{
  items: AdoWorkItemDto[]
  masterList: MasterListState
  added: number
  updated: number
  removed: number
  errors: string[]
  syncedAs: string[]
}> {
  const connections = getConnections().filter(
    (c) => c.active && c.organization && c.pat
  )
  if (connections.length === 0) {
    throw new Error('Turn on at least one organization in Settings.')
  }

  const all: WorkItemWithConn[] = []
  const errors: string[] = []
  const syncedAs: string[] = []

  for (const conn of connections) {
    try {
      const { items, asUser } = await fetchMyWorkItems(credentialsFor(conn))
      for (const item of items) {
        all.push({ ...item, connectionId: conn.id })
      }
      const who = asUser.uniqueName
        ? `${asUser.displayName} <${asUser.uniqueName}>`
        : asUser.displayName
      syncedAs.push(`${conn.organization}: @Me = ${who}`)
    } catch (e) {
      errors.push(
        `${conn.organization}: ${e instanceof Error ? e.message : 'Sync failed'}`
      )
    }
  }

  if (all.length === 0 && errors.length > 0) {
    throw new Error(errors.join('\n'))
  }

  const { list, added, updated, removed, newTickets, hadPriorList } =
    replaceMasterFromSync(all)
  if (hadPriorList && newTickets.length > 0) {
    alertNewTickets(newTickets)
  }
  return {
    items: all.map(toDto),
    masterList: list,
    added,
    updated,
    removed,
    errors,
    syncedAs
  }
}

function setupAdoAuthenticatedMedia(): void {
  const filter = {
    urls: [
      'https://dev.azure.com/*',
      'https://*.dev.azure.com/*',
      'https://*.visualstudio.com/*',
      'https://*.vsassets.io/*',
      'https://*.vssps.visualstudio.com/*'
    ]
  }

  session.defaultSession.webRequest.onBeforeSendHeaders(filter, (details, callback) => {
    const requestHeaders = { ...details.requestHeaders }
    try {
      const url = new URL(details.url)
      let org = ''
      if (url.hostname === 'dev.azure.com' || url.hostname.endsWith('.dev.azure.com')) {
        org = decodeURIComponent(url.pathname.split('/').filter(Boolean)[0] || '')
      } else if (url.hostname.endsWith('.visualstudio.com')) {
        org = url.hostname.split('.')[0] || ''
      }

      const connections = getConnections()
      const conn =
        connections.find(
          (c) => c.organization.toLowerCase() === org.toLowerCase() && c.pat
        ) || connections.find((c) => c.pat)

      if (conn?.pat) {
        requestHeaders.Authorization = `Basic ${Buffer.from(`:${conn.pat}`).toString('base64')}`
      }
    } catch {
      // leave headers unchanged
    }
    callback({ requestHeaders })
  })
}

function isAdoMediaHost(hostname: string): boolean {
  const host = hostname.toLowerCase()
  return (
    host === 'dev.azure.com' ||
    host.endsWith('.dev.azure.com') ||
    host.endsWith('.visualstudio.com') ||
    host.endsWith('.vsassets.io') ||
    host.endsWith('.vssps.visualstudio.com') ||
    host.endsWith('.azureedge.net')
  )
}

function patForMediaUrl(url: URL): string | undefined {
  let org = ''
  if (url.hostname === 'dev.azure.com' || url.hostname.endsWith('.dev.azure.com')) {
    org = decodeURIComponent(url.pathname.split('/').filter(Boolean)[0] || '')
  } else if (url.hostname.endsWith('.visualstudio.com')) {
    org = url.hostname.split('.')[0] || ''
  }
  const connections = getConnections()
  const conn =
    connections.find(
      (c) => c.organization.toLowerCase() === org.toLowerCase() && c.pat
    ) || connections.find((c) => c.pat)
  return conn?.pat
}

function setupAdoImageProtocol(): void {
  protocol.handle('adoimg', async (request) => {
    try {
      const parsed = new URL(request.url)
      const target = parsed.searchParams.get('u')
      if (!target) {
        return new Response('Missing image URL', { status: 400 })
      }
      const dest = new URL(target)
      if (!isAdoMediaHost(dest.hostname)) {
        return new Response('Blocked host', { status: 403 })
      }
      if (!dest.searchParams.has('api-version') && dest.pathname.includes('/_apis/wit/attachments/')) {
        dest.searchParams.set('api-version', '7.1')
      }
      const headers: Record<string, string> = {
        Accept: '*/*'
      }
      const pat = patForMediaUrl(dest)
      if (pat) {
        headers.Authorization = `Basic ${Buffer.from(`:${pat}`).toString('base64')}`
      }
      const res = await net.fetch(dest.toString(), { headers, redirect: 'follow' })
      return res
    } catch (err) {
      console.error('adoimg proxy failed', err)
      return new Response('Image proxy failed', { status: 502 })
    }
  })
}

function playAlertSound(soundName?: string): void {
  const name = normalizeAlertSound(
    soundName ??
      normalizeSettings(store.get('settings') as Record<string, unknown>).alertSound
  )
  playNamedAlert(name)
}

function alertNewTickets(tickets: MasterEntry[]): void {
  if (tickets.length === 0) return

  const first = tickets[0]
  const title =
    tickets.length === 1
      ? `New ticket #${first.workItemId}`
      : `${tickets.length} new tickets`
  const body =
    tickets.length === 1
      ? `${first.organization}: ${first.title}`
      : tickets
          .slice(0, 3)
          .map((t) => `#${t.workItemId} ${t.title}`)
          .join('\n')

  if (Notification.isSupported()) {
    const note = new Notification({
      title,
      body,
      silent: true
    })
    note.on('click', () => {
      if (!mainWindow) return
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.show()
      mainWindow.focus()
    })
    note.show()
  }

  if (process.platform === 'darwin') {
    playAlertSound()
    app.dock?.bounce('critical')
  }
}

let outlookTimer: ReturnType<typeof setInterval> | null = null
let outlookSnapshot: OutlookSnapshot = emptyOutlookSnapshot()
let outlookRefreshInFlight: Promise<OutlookTokens | null> | null = null

function pushOutlookSnapshot(next: OutlookSnapshot): void {
  outlookSnapshot = next
  mainWindow?.webContents.send('outlook:snapshot', next)
}

async function validOutlookTokens(): Promise<OutlookTokens | null> {
  const tokens = store.get('outlookAuth')
  if (!tokens?.accessToken) return null
  if (tokens.expiresAt > Date.now() + 15_000) return tokens
  if (!outlookRefreshInFlight) {
    outlookRefreshInFlight = refreshOutlookTokens(tokens)
      .then((next) => {
        store.set('outlookAuth', next)
        return next
      })
      .catch((err) => {
        console.error('outlook refresh failed', err)
        store.set('outlookAuth', null)
        pushOutlookSnapshot(
          emptyOutlookSnapshot(
            err instanceof Error ? err.message : 'Outlook session expired. Sign in again.'
          )
        )
        return null
      })
      .finally(() => {
        outlookRefreshInFlight = null
      })
  }
  return outlookRefreshInFlight
}

function notifyOutlook(title: string, body: string, url?: string): void {
  if (Notification.isSupported()) {
    const note = new Notification({
      title,
      body,
      silent: true
    })
    note.on('click', () => {
      if (url) void shell.openExternal(url)
      if (!mainWindow) return
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.show()
      mainWindow.focus()
    })
    note.show()
  }
  if (process.platform === 'darwin') {
    playAlertSound()
    app.dock?.bounce('critical')
  }
}

async function pollOutlook(): Promise<void> {
  const settings = normalizeSettings(store.get('settings') as Record<string, unknown>)
  const tokens = await validOutlookTokens()
  if (!tokens) {
    if (outlookSnapshot.connected) pushOutlookSnapshot(emptyOutlookSnapshot())
    return
  }

  try {
    const data = await loadOutlookSnapshot(tokens.accessToken, {
      calendar: settings.outlookCalendar,
      mail: settings.outlookMailAlerts
    })
    const seen = store.get('outlookSeen') ?? {
      mailIds: [],
      meetingKeys: [],
      seeded: false
    }
    const knownMail = new Set(seen.mailIds)
    const knownMeetings = new Set(seen.meetingKeys)

    if (settings.outlookMailAlerts && seen.seeded) {
      const unreadNew = data.mail.filter((m) => !m.isRead && !knownMail.has(m.id))
      if (unreadNew.length === 1) {
        notifyOutlook(
          `New mail: ${unreadNew[0].subject}`,
          unreadNew[0].from,
          unreadNew[0].webLink
        )
      } else if (unreadNew.length > 1) {
        notifyOutlook(
          `${unreadNew.length} new Outlook messages`,
          unreadNew
            .slice(0, 3)
            .map((m) => m.subject)
            .join('\n')
        )
      }
    }

    if (settings.outlookCalendar) {
      const leadMs = settings.outlookMeetingAlertMinutes * 60 * 1000
      const now = Date.now()
      for (const ev of data.events) {
        if (ev.isAllDay) continue
        const start = new Date(ev.start.endsWith('Z') ? ev.start : `${ev.start}Z`).getTime()
        if (Number.isNaN(start)) continue
        const key = `${ev.id}|${ev.start}`
        const inWindow = now >= start - leadMs && now <= start + 2 * 60 * 1000
        if (inWindow && !knownMeetings.has(key)) {
          if (seen.seeded) {
            const when =
              settings.outlookMeetingAlertMinutes === 0
                ? 'starting now'
                : `in ${settings.outlookMeetingAlertMinutes} min`
            notifyOutlook(`Meeting ${when}`, ev.subject, ev.webLink)
          }
          knownMeetings.add(key)
        }
      }
    }

    store.set('outlookSeen', {
      mailIds: [...new Set([...seen.mailIds, ...data.mail.map((m) => m.id)])].slice(-500),
      meetingKeys: [...knownMeetings].slice(-500),
      seeded: true
    })

    pushOutlookSnapshot({
      connected: true,
      account: tokens.account,
      error: null,
      events: settings.outlookCalendar ? data.events : [],
      mail: settings.outlookMailAlerts ? data.mail.filter((m) => !m.isRead) : [],
      unreadCount: data.unreadCount
    })
  } catch (err) {
    const status = (err as { status?: number }).status
    if (status === 401) store.set('outlookAuth', null)
    pushOutlookSnapshot({
      connected: Boolean(store.get('outlookAuth')),
      account: store.get('outlookAuth')?.account || '',
      error: err instanceof Error ? err.message : 'Outlook sync failed',
      events: [],
      mail: [],
      unreadCount: 0
    })
  }
}

function startOutlookPolling(): Promise<void> {
  if (outlookTimer) clearInterval(outlookTimer)
  outlookTimer = setInterval(() => {
    void pollOutlook()
  }, 60_000)
  return pollOutlook()
}

function createWindow(): void {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 960,
    minHeight: 640,
    title: 'ADO Master List',
    backgroundColor: '#e8eef5',
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      webSecurity: true
    }
  })
  mainWindow = win
  win.on('closed', () => {
    if (mainWindow === win) mainWindow = null
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    win.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    win.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

app.whenReady().then(() => {
  process.on('uncaughtException', (err) => {
    console.error('uncaughtException', err)
  })
  process.on('unhandledRejection', (err) => {
    console.error('unhandledRejection', err)
  })
  setupAdoAuthenticatedMedia()
  setupAdoImageProtocol()
  createWindow()
  startOutlookPolling()
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  // Keep running on macOS; only quit on other platforms
  if (process.platform !== 'darwin') app.quit()
})

ipcMain.handle('settings:get', () =>
  normalizeSettings(store.get('settings') as Record<string, unknown>)
)

ipcMain.handle('settings:save', (_e, settings: AppSettings) => {
  const normalized = normalizeSettings(settings)
  store.set('settings', normalized)
  startOutlookPolling()
  return store.get('settings')
})

function readMasterList(): MasterListState {
  const current = store.get('masterList')
  const entries = (current.entries ?? []).filter(
    (e) => !isExcludedWorkItemState(e.state)
  )
  if (entries.length !== (current.entries ?? []).length) {
    const next = { entries }
    store.set('masterList', next)
    return next
  }
  return { entries }
}

ipcMain.handle('master:get', () => readMasterList())

ipcMain.handle('master:save', (_e, state: MasterListState) => {
  const next = {
    entries: state.entries.filter((e) => !isExcludedWorkItemState(e.state))
  }
  store.set('masterList', next)
  return next
})

ipcMain.handle('ado:listProjects', async (_e, connectionId: string): Promise<AdoProjectDto[]> => {
  const conn = getConnection(connectionId)
  const projects = await fetchProjects(credentialsFor(conn))
  return projects.map((p) => ({
    ...p,
    organization: conn.organization,
    connectionId: conn.id
  }))
})

ipcMain.handle('ado:syncMyWorkItems', async () => syncAllConnections())

ipcMain.handle(
  'ado:getWorkItemComments',
  async (
    _e,
    payload: { connectionId: string; id: number; project: string }
  ) => {
    const conn = getConnection(payload.connectionId)
    return fetchWorkItemComments(
      credentialsFor(conn),
      payload.id,
      payload.project
    )
  }
)

ipcMain.handle(
  'ado:getWorkItem',
  async (_e, payload: { connectionId: string; id: number }) => {
    const conn = getConnection(payload.connectionId)
    const item = await fetchWorkItem(credentialsFor(conn), payload.id)
    return toDto({ ...item, connectionId: conn.id })
  }
)

ipcMain.handle(
  'ado:updateWorkItem',
  async (
    _e,
    payload: {
      connectionId: string
      id: number
      title?: string
      state?: string
      description?: string
    }
  ): Promise<AdoWorkItemDto> => {
    const conn = getConnection(payload.connectionId)
    const item = await updateWorkItem(credentialsFor(conn), payload.id, {
      title: payload.title,
      state: payload.state,
      description: payload.description
    })
    return toDto({ ...item, connectionId: conn.id })
  }
)

ipcMain.handle('shell:openExternal', async (_e, url: string) => {
  await shell.openExternal(url)
})

ipcMain.handle('sound:play', (_e, soundName?: string) => {
  playAlertSound(soundName)
})

ipcMain.handle('outlook:get', () => outlookSnapshot)

ipcMain.handle('outlook:connect', async (_e, clientId: string) => {
  const id = String(clientId ?? '').trim()
  const current = normalizeSettings(store.get('settings') as Record<string, unknown>)
  store.set('settings', { ...current, outlookClientId: id || current.outlookClientId })
  const tokens = await signInToOutlook(id || current.outlookClientId, mainWindow)
  store.set('outlookAuth', tokens)
  store.set('outlookSeen', { mailIds: [], meetingKeys: [], seeded: false })
  await startOutlookPolling()
  return outlookSnapshot
})

ipcMain.handle('outlook:disconnect', () => {
  store.set('outlookAuth', null)
  store.set('outlookSeen', { mailIds: [], meetingKeys: [], seeded: false })
  pushOutlookSnapshot(emptyOutlookSnapshot())
  return outlookSnapshot
})

ipcMain.handle('master:updateEntry', (_e, entry: MasterEntry) => {
  const list = store.get('masterList')
  const next = {
    entries: list.entries.map((e) => (e.id === entry.id ? entry : e))
  }
  store.set('masterList', next)
  return next
})
