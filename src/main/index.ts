import { app, BrowserWindow, ipcMain, shell, session, Notification, protocol, net } from 'electron'
import { decryptSecret, encryptSecret, encryptionAvailable } from './safeSecrets'
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
import { attachEditContextMenu, installAppMenu } from './editMenu'
import {
  disposeMsWebEmbeds,
  hideMsWebEmbed,
  reloadMsWebEmbed,
  showMsWebEmbed,
  focusMsWebEmbed,
  setMsWebNotifyHandler,
  setMsWebCountHandler,
  startMsWebCountPolling,
  getMsWebCounts,
  syncMsWebCounts,
  pruneMsWebEmbeds,
  type MsWebEmbedId
} from './msWebEmbed'
import {
  emptyOutlookSnapshot,
  loadOutlookDay,
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
  OutlookSnapshot,
  CustomWebTab,
  LocalNote,
  NowPin
} from '../../shared/types'
import { normalizeAlertSound, normalizeCustomWebTab, normalizeLocalNotes, normalizeNowPins } from '../../shared/types'

type WorkItemWithConn = AdoWorkItem & { connectionId: string }

if (!app.isReady()) {
  app.commandLine.appendSwitch(
    'enable-features',
    'MacSckSystemPicker,MacLoopbackAudioForScreenShare'
  )
}

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

app.commandLine.appendSwitch('autoplay-policy', 'no-user-gesture-required')

let mainWindow: BrowserWindow | null = null

const store = new Store<{
  settings: AppSettings & Record<string, unknown>
  masterList: MasterListState
  outlookAuth: OutlookTokens | string | null
  outlookSeen: { mailIds: string[]; meetingKeys: string[]; seeded: boolean }
  localNotes: LocalNote[]
  nowPins: NowPin[]
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
      outlookMeetingAlertMinutes: 5,
      webActivityAlerts: true,
      outlookAlertSound: 'VoiceMailAlarm',
      teamsAlertSound: 'VoiceTeamsAlarm',
      customTabs: []
    },
    masterList: {
      entries: []
    },
    outlookAuth: null,
    outlookSeen: { mailIds: [], meetingKeys: [], seeded: false },
    localNotes: [],
    nowPins: []
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
  const allowed = new Set([0, 2, 5, 15, 30, 60])
  const meetMins = Number(raw.outlookMeetingAlertMinutes)
  const meetAllowed = new Set([0, 5, 10, 15])

  return {
    connections,
    autoSyncMinutes: allowed.has(minutes) ? minutes : 0,
    alertSound: normalizeAlertSound(raw.alertSound),
    outlookClientId: String(raw.outlookClientId ?? '').trim(),
    outlookCalendar: raw.outlookCalendar !== false,
    outlookMailAlerts: raw.outlookMailAlerts !== false,
    outlookMeetingAlertMinutes: meetAllowed.has(meetMins) ? meetMins : 5,
    webActivityAlerts: raw.webActivityAlerts !== false,
    outlookAlertSound: normalizeAlertSound(raw.outlookAlertSound ?? raw.alertSound),
    teamsAlertSound: normalizeAlertSound(raw.teamsAlertSound ?? raw.alertSound),
    customTabs: Array.isArray(raw.customTabs)
      ? (raw.customTabs as Partial<CustomWebTab>[])
          .map(normalizeCustomWebTab)
          .filter((t): t is CustomWebTab => Boolean(t))
      : []
  }
}

function normalizeSettings(settings: AppSettings | Record<string, unknown>): AppSettings {
  return migrateLegacySettings(settings as Record<string, unknown>)
}

function decryptSettings(raw: Record<string, unknown>): AppSettings {
  const normalized = normalizeSettings(raw)
  return {
    ...normalized,
    connections: normalized.connections.map((c) => ({
      ...c,
      pat: decryptSecret(c.pat)
    }))
  }
}

function readSettings(): AppSettings {
  return decryptSettings(store.get('settings') as Record<string, unknown>)
}

function writeSettings(settings: AppSettings): AppSettings {
  const normalized = normalizeSettings(settings)
  store.set('settings', {
    ...normalized,
    connections: normalized.connections.map((c) => ({
      ...c,
      pat: encryptSecret(c.pat)
    }))
  })
  return normalized
}

function readOutlookAuth(): OutlookTokens | null {
  const raw = store.get('outlookAuth')
  if (!raw) return null
  if (typeof raw === 'string') {
    const json = decryptSecret(raw)
    if (!json) return null
    try {
      return JSON.parse(json) as OutlookTokens
    } catch {
      return null
    }
  }
  if (typeof raw === 'object' && raw && 'accessToken' in raw) {
    return raw as OutlookTokens
  }
  return null
}

function writeOutlookAuth(tokens: OutlookTokens | null): void {
  if (!tokens) {
    store.set('outlookAuth', null)
    return
  }
  if (encryptionAvailable()) {
    store.set('outlookAuth', encryptSecret(JSON.stringify(tokens)))
    return
  }
  store.set('outlookAuth', tokens)
}

function importLegacyOrgStore(): void {
  try {
    const legacy = new Store({ name: 'config' })
    const legacySettings = legacy.get('settings') as Record<string, unknown> | undefined
    const current = readSettings()
    if (
      legacySettings &&
      current.connections.length === 0 &&
      (legacySettings.organization || legacySettings.pat)
    ) {
      writeSettings(
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
}

function sealSecretsAtRest(): void {
  importLegacyOrgStore()
  try {
    writeSettings(readSettings())
    const outlook = readOutlookAuth()
    if (outlook) writeOutlookAuth(outlook)
  } catch (err) {
    console.error('Failed to encrypt secrets at rest', err)
  }
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

/** Keep existing tickets; update from this sync; drop closed/unassigned for orgs that synced. */
function mergeMasterFromSync(
  items: WorkItemWithConn[],
  syncedOrganizations: string[]
): {
  list: MasterListState
  added: number
  updated: number
  removed: number
  newTickets: MasterEntry[]
  hadPriorList: boolean
} {
  const existing = store.get('masterList').entries
  const now = new Date().toISOString()
  const synced = new Set(
    syncedOrganizations.map((org) => org.trim().toLowerCase()).filter(Boolean)
  )

  const incoming = new Map<string, WorkItemWithConn>()
  for (const item of items) {
    if (isExcludedWorkItemState(item.state)) continue
    const id = masterEntryId(item.organization, item.id)
    if (!incoming.has(id)) incoming.set(id, item)
  }

  const next: MasterEntry[] = []
  const newTickets: MasterEntry[] = []
  const keptIncoming = new Set<string>()
  let added = 0
  let updated = 0
  let removed = 0

  for (const prev of existing) {
    const orgKey = (prev.organization || '').trim().toLowerCase()
    if (!synced.has(orgKey)) {
      next.push(prev)
      continue
    }
    const id = prev.id || masterEntryId(prev.organization, prev.workItemId)
    const item = incoming.get(id)
    if (!item) {
      removed += 1
      continue
    }
    keptIncoming.add(id)
    next.push(entryFromWorkItem(item, prev, now))
    updated += 1
  }

  for (const [id, item] of incoming) {
    if (keptIncoming.has(id)) continue
    const entry = entryFromWorkItem(item, undefined, now)
    added += 1
    newTickets.push(entry)
    next.push(entry)
  }

  const list = { entries: next }
  store.set('masterList', list)

  const keptWorkIds = new Set(next.map((e) => e.id))
  const pins = normalizeNowPins(store.get('nowPins')).filter(
    (p) => p.kind !== 'work' || keptWorkIds.has(p.id)
  )
  store.set('nowPins', pins)

  return {
    list,
    added,
    updated,
    removed,
    newTickets,
    hadPriorList: existing.length > 0
  }
}

function getConnections(): AdoOrgConnection[] {
  return readSettings().connections
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
  const syncedOrganizations: string[] = []

  for (const conn of connections) {
    try {
      const { items, asUser } = await fetchMyWorkItems(credentialsFor(conn))
      for (const item of items) {
        all.push({ ...item, connectionId: conn.id })
      }
      syncedOrganizations.push(conn.organization)
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

  if (all.length === 0 && errors.length > 0 && syncedOrganizations.length === 0) {
    throw new Error(errors.join('\n'))
  }

  const { list, added, updated, removed, newTickets, hadPriorList } =
    mergeMasterFromSync(all, syncedOrganizations)
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
      readSettings().alertSound
  )
  playNamedAlert(name)
}

let lastWebAlertKey = ''
let lastWebAlertAt = 0
let lastWebAlertSource = ''
let alertsMuted = false
let lastMsCounts = { outlookUnread: -1, teamsUnread: -1 }

function setAlertsMuted(muted: boolean): boolean {
  alertsMuted = Boolean(muted)
  mainWindow?.webContents.send('alerts:muted', alertsMuted)
  return alertsMuted
}

function alertAttention(
  title: string,
  body: string,
  onClick?: () => void,
  soundName?: string
): void {
  if (alertsMuted) return
  if (Notification.isSupported()) {
    const note = new Notification({
      title,
      body,
      silent: true
    })
    note.on('click', () => {
      onClick?.()
      if (!mainWindow) return
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.show()
      mainWindow.focus()
    })
    note.show()
  }
  if (process.platform === 'darwin') {
    playAlertSound(soundName)
    app.dock?.bounce('critical')
  }
}

function handleMsWebNotify(payload: { source: MsWebEmbedId; title: string; body: string }): void {
  if (alertsMuted) return
  if (payload.source !== 'outlook' && payload.source !== 'teams') return
  const settings = readSettings()
  if (!settings.webActivityAlerts) return
  const soundName =
    payload.source === 'teams' ? settings.teamsAlertSound : settings.outlookAlertSound
  if (soundName === 'none') return
  const title = payload.title.trim()
  const body = payload.body.trim()
  if (!title && !body) return
  const blob = `${title} ${body}`.toLowerCase()
  if (blob.includes('is typing') || blob.includes('are typing')) return
  if (payload.source === 'outlook') {
    if (
      /meeting|calendar|event reminder|starts in|starting now|canceled|cancelled|invitation|invite to|accepted:|declined:|tentative|new unread mail/.test(
        blob
      )
    ) {
      return
    }
  }
  const key = `${payload.source}|${title}|${body}`
  const now = Date.now()
  if (key === lastWebAlertKey && now - lastWebAlertAt < 5000) return
  if (payload.source === lastWebAlertSource && now - lastWebAlertAt < 8000) return
  lastWebAlertKey = key
  lastWebAlertAt = now
  lastWebAlertSource = payload.source
  const prefix = payload.source === 'teams' ? 'Teams' : 'Outlook'
  alertAttention(
    title || prefix,
    body || `New ${prefix} notification`,
    () => {
      mainWindow?.webContents.send('ms-web:open-tab', payload.source)
    },
    soundName
  )
}

function alertNewTickets(tickets: MasterEntry[]): void {
  if (alertsMuted || tickets.length === 0) return

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
  const tokens = readOutlookAuth()
  if (!tokens?.accessToken) return null
  if (tokens.expiresAt > Date.now() + 15_000) return tokens
  if (!outlookRefreshInFlight) {
    outlookRefreshInFlight = refreshOutlookTokens(tokens)
      .then((next) => {
        writeOutlookAuth(next)
        return next
      })
      .catch((err) => {
        console.error('outlook refresh failed', err)
        writeOutlookAuth(null)
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
  if (alertsMuted) return
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
  const settings = readSettings()
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
    if (status === 401) writeOutlookAuth(null)
    const remaining = readOutlookAuth()
    pushOutlookSnapshot({
      connected: Boolean(remaining),
      account: remaining?.account || '',
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
  attachEditContextMenu(win.webContents)
  win.on('closed', () => {
    if (mainWindow === win) {
      disposeMsWebEmbeds()
      mainWindow = null
    }
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
  installAppMenu()
  sealSecretsAtRest()
  setMsWebNotifyHandler(handleMsWebNotify)
  setMsWebCountHandler((next) => {
    mainWindow?.webContents.send('ms-web:counts', next)
    if (lastMsCounts.teamsUnread > 0 && next.teamsUnread > lastMsCounts.teamsUnread) {
      handleMsWebNotify({
        source: 'teams',
        title: 'Teams',
        body: 'New unread message'
      })
    }
    lastMsCounts = {
      outlookUnread: next.outlookUnread,
      teamsUnread: next.teamsUnread
    }
  })
  startMsWebCountPolling()
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

ipcMain.handle('notes:get', () => normalizeLocalNotes(store.get('localNotes')))

ipcMain.handle('notes:save', (_e, notes: LocalNote[]) => {
  const next = normalizeLocalNotes(notes)
  store.set('localNotes', next)
  return next
})

ipcMain.handle('now:get', () => normalizeNowPins(store.get('nowPins')))

ipcMain.handle('now:save', (_e, pins: NowPin[]) => {
  const next = normalizeNowPins(pins)
  store.set('nowPins', next)
  return next
})

ipcMain.handle('settings:get', () => readSettings())

ipcMain.handle('settings:save', (_e, settings: AppSettings) => {
  const normalized = writeSettings(settings)
  pruneMsWebEmbeds(normalized.customTabs.map((t) => t.id))
  startOutlookPolling()
  return normalized
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

ipcMain.handle('alerts:getMuted', () => alertsMuted)

ipcMain.handle('alerts:setMuted', (_e, muted: boolean) => setAlertsMuted(muted))

ipcMain.handle(
  'msWeb:show',
  (
    _e,
    payload: {
      id: string
      url?: string
      bounds: { x: number; y: number; width: number; height: number }
    }
  ) => {
    if (!mainWindow || !payload?.id) return
    showMsWebEmbed(mainWindow, payload.id, payload.bounds, payload.url)
  }
)

ipcMain.handle('msWeb:hide', () => {
  hideMsWebEmbed(mainWindow)
})

ipcMain.handle('msWeb:focus', () => {
  focusMsWebEmbed()
})

ipcMain.handle('msWeb:reload', (_e, payload: { id: string; url?: string } | string) => {
  if (typeof payload === 'string') {
    reloadMsWebEmbed(payload)
    return
  }
  if (payload?.id) reloadMsWebEmbed(payload.id, payload.url)
})

ipcMain.handle('msWeb:counts', () => getMsWebCounts())

ipcMain.handle('msWeb:syncCounts', () => syncMsWebCounts())

ipcMain.handle('outlook:get', () => outlookSnapshot)

ipcMain.handle('outlook:day', async (_e, ymd: string) => {
  const tokens = await validOutlookTokens()
  if (!tokens) throw new Error('Sign in to Outlook in Settings.')
  return loadOutlookDay(tokens.accessToken, String(ymd ?? ''))
})

ipcMain.handle('outlook:connect', async (_e, clientId: string) => {
  const id = String(clientId ?? '').trim()
  const current = readSettings()
  writeSettings({ ...current, outlookClientId: id || current.outlookClientId })
  const tokens = await signInToOutlook(id || current.outlookClientId, mainWindow)
  writeOutlookAuth(tokens)
  store.set('outlookSeen', { mailIds: [], meetingKeys: [], seeded: false })
  await startOutlookPolling()
  return outlookSnapshot
})

ipcMain.handle('outlook:disconnect', () => {
  writeOutlookAuth(null)
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
