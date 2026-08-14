import { useCallback, useEffect, useMemo, useState } from 'react'
import {
  ALERT_SOUNDS,
  ALERT_SOUND_LABELS,
  type AdoCommentDto,
  type AdoOrgConnection,
  type AppSettings,
  type CustomWebTab,
  type MasterEntry,
  type NowPin,
  type OutlookSnapshot,
  type WorkStatus,
  NOW_PIN_LIMIT,
  nowPinKey,
  toggleNowPin
} from '../../shared/types'
import { RichHtml } from './html'
import { MsWebEmbed } from './MsWebEmbed'
import { NotesPane } from './NotesPane'
import { NowPane } from './NowPane'

type ModalMode = 'settings' | null

const emptyOutlook = (): OutlookSnapshot => ({
  connected: false,
  account: '',
  error: null,
  events: [],
  mail: [],
  unreadCount: 0
})

const emptySettings = (): AppSettings => ({
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
})

function newLocalId(): string {
  return `conn-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function emptyCustomTab(): CustomWebTab {
  return {
    id: `tab-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
    label: '',
    url: ''
  }
}

function emptyConnection(): AdoOrgConnection {
  return {
    id: newLocalId(),
    organization: '',
    pat: '',
    enabledProjects: [],
    active: true
  }
}

function ticketAgeClass(entry: MasterEntry): string {
  const raw = entry.createdDate || entry.changedDate || entry.addedAt
  if (!raw) return ''
  const created = new Date(raw).getTime()
  if (Number.isNaN(created)) return ''
  const hours = (Date.now() - created) / 3_600_000
  if (hours <= 1) return 'age-1h'
  if (hours <= 4) return 'age-4h'
  if (hours <= 6) return 'age-6h'
  return 'age-8h'
}

function statusClass(status: WorkStatus): string {
  if (status === 'active') return 'chip chip-active'
  if (status === 'blocked') return 'chip chip-blocked'
  if (status === 'done') return 'chip chip-done'
  return 'chip chip-warn'
}

function formatMeetingWhen(start: string, end: string, isAllDay: boolean): string {
  if (isAllDay) return 'All day'
  const s = new Date(start.endsWith('Z') ? start : `${start}Z`)
  const e = new Date(end.endsWith('Z') ? end : `${end}Z`)
  if (Number.isNaN(s.getTime())) return ''
  const startLabel = s.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })
  if (Number.isNaN(e.getTime())) return startLabel
  const endLabel = e.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })
  return `${startLabel}–${endLabel}`
}

function formatDate(value?: string): string {
  if (!value) return '—'
  const d = new Date(value)
  if (Number.isNaN(d.getTime())) return value
  return d.toLocaleString(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  })
}

function normalizeAppSettings(s: AppSettings): AppSettings {
  return {
    connections: (Array.isArray(s.connections) ? s.connections : []).map((c) => ({
      ...c,
      enabledProjects: Array.isArray(c.enabledProjects) ? c.enabledProjects : [],
      active: c.active !== false
    })),
    autoSyncMinutes: s.autoSyncMinutes ?? 0,
    alertSound: s.alertSound || 'Alarm',
    outlookClientId: s.outlookClientId ?? '',
    outlookCalendar: s.outlookCalendar !== false,
    outlookMailAlerts: s.outlookMailAlerts !== false,
    outlookMeetingAlertMinutes: [0, 5, 10, 15].includes(s.outlookMeetingAlertMinutes)
      ? s.outlookMeetingAlertMinutes
      : 5,
    webActivityAlerts: s.webActivityAlerts !== false,
    outlookAlertSound: s.outlookAlertSound || 'VoiceMailAlarm',
    teamsAlertSound: s.teamsAlertSound || 'VoiceTeamsAlarm',
    customTabs: Array.isArray(s.customTabs) ? s.customTabs : []
  }
}

function AlertSoundPicker(props: {
  id: string
  label: string
  value: string
  onChange: (value: string) => void
}) {
  const value = props.value || 'Alarm'
  return (
    <div className="field" style={{ marginBottom: '0.5rem' }}>
      <label htmlFor={props.id}>{props.label}</label>
      <div className="setup-sound-row">
        <select
          id={props.id}
          value={value}
          onChange={(e) => props.onChange(e.target.value)}
        >
          {ALERT_SOUNDS.map((name) => (
            <option key={name} value={name}>
              {ALERT_SOUND_LABELS[name] ?? name}
            </option>
          ))}
        </select>
        <button
          className="btn"
          type="button"
          disabled={value === 'none'}
          onClick={() => void window.adoApi.playAlertSound(value)}
        >
          Test
        </button>
      </div>
    </div>
  )
}

type ThemeMode = 'light' | 'dark'

function readStoredTheme(): ThemeMode {
  try {
    const stored = localStorage.getItem('ado-theme')
    if (stored === 'dark' || stored === 'light') return stored
  } catch {
    // ignore
  }
  if (window.matchMedia?.('(prefers-color-scheme: dark)').matches) return 'dark'
  return 'light'
}

function applyTheme(theme: ThemeMode) {
  document.documentElement.dataset.theme = theme
}

export default function App() {
  const [settings, setSettings] = useState<AppSettings>(emptySettings)
  const [master, setMaster] = useState<MasterEntry[]>([])
  const [search, setSearch] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [notice, setNotice] = useState<string | null>(null)
  const [modal, setModal] = useState<ModalMode>(null)
  const [draftSettings, setDraftSettings] = useState<AppSettings>(emptySettings)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [lastSyncAt, setLastSyncAt] = useState<string | null>(null)
  const [comments, setComments] = useState<AdoCommentDto[]>([])
  const [commentsLoading, setCommentsLoading] = useState(false)
  const [commentsError, setCommentsError] = useState<string | null>(null)
  const [theme, setTheme] = useState<ThemeMode>(() => {
    const initial = readStoredTheme()
    applyTheme(initial)
    return initial
  })
  const [expandedOrgs, setExpandedOrgs] = useState<Set<string>>(() => new Set())
  const [outlook, setOutlook] = useState<OutlookSnapshot>(emptyOutlook)
  const [screen, setScreen] = useState<'work' | 'calendar' | 'teams' | 'notes' | 'now' | string>(
    'now'
  )
  const [notesOpenCount, setNotesOpenCount] = useState(0)
  const [nowPins, setNowPins] = useState<NowPin[]>([])
  const [focusNoteId, setFocusNoteId] = useState<string | null>(null)
  const [notesEpoch, setNotesEpoch] = useState(0)
  const [webCounts, setWebCounts] = useState({
    outlookUnread: 0,
    teamsUnread: 0,
    extra: {} as Record<string, number>
  })
  const [alertsMuted, setAlertsMuted] = useState(false)

  useEffect(() => {
    applyTheme(theme)
    try {
      localStorage.setItem('ado-theme', theme)
    } catch {
      // ignore
    }
  }, [theme])

  const hasConnection = settings.connections.some(
    (c) => c.active !== false && c.organization && c.pat
  )

  const syncMine = useCallback(async (opts?: { quiet?: boolean }) => {
    setLoading(true)
    if (!opts?.quiet) {
      setError(null)
      setNotice(null)
    }
    try {
      const countsPromise = window.adoApi.syncMsWebCounts()
      if (hasConnection) {
        const result = await window.adoApi.syncMyWorkItems()
        setMaster(result.masterList.entries)
        void window.adoApi.getNowPins().then(setNowPins)
        setLastSyncAt(new Date().toISOString())
        setSelectedId((current) => {
          if (current && result.masterList.entries.some((e) => e.id === current)) {
            return current
          }
          return result.masterList.entries[0]?.id ?? null
        })
        const errSuffix =
          result.errors?.length > 0 ? ` · ${result.errors.length} org error(s)` : ''
        const asMe =
          result.syncedAs?.length > 0 ? ` · ${result.syncedAs.join(' · ')}` : ''
        setNotice(
          `Synced @Me · +${result.added} new · −${result.removed} closed · ${result.updated} updated${errSuffix}${asMe}`
        )
        if (result.errors?.length) {
          setError(result.errors.join('\n'))
        }
      } else {
        setLastSyncAt(new Date().toISOString())
        setNotice('Synced Outlook and Teams counts')
      }
      try {
        setWebCounts(await countsPromise)
      } catch {
        // badges still update via the 8s poll
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to load work items')
    } finally {
      setLoading(false)
    }
  }, [hasConnection])

  const loadLocal = useCallback(async () => {
    const [s, list] = await Promise.all([
      window.adoApi.getSettings(),
      window.adoApi.getMasterList()
    ])
    const normalized = normalizeAppSettings(s)
    setSettings(normalized)
    setDraftSettings(normalized)
    setMaster(list.entries)
    if (list.entries[0]) setSelectedId(list.entries[0].id)
    if (!normalized.connections.some((c) => c.active !== false && c.organization && c.pat)) {
      setModal('settings')
      return
    }
    // Always refresh latest assigned items on launch
    void syncMine({ quiet: true })
  }, [syncMine])

  useEffect(() => {
    void loadLocal()
  }, [loadLocal])

  useEffect(() => {
    void window.adoApi.getOutlook().then(setOutlook)
    return window.adoApi.onOutlook(setOutlook)
  }, [])

  useEffect(() => {
    void window.adoApi.getLocalNotes().then((list) =>
      setNotesOpenCount(list.filter((n) => !n.done).length)
    )
    void window.adoApi.getNowPins().then(setNowPins)
  }, [])

  useEffect(() => {
    void window.adoApi.getAlertsMuted().then(setAlertsMuted)
    return window.adoApi.onAlertsMuted(setAlertsMuted)
  }, [])

  useEffect(() => {
    return window.adoApi.onMsWebOpenTab((id) => {
      if (id === 'teams') setScreen('teams')
      else if (id === 'outlook') setScreen('calendar')
      else setScreen(id)
    })
  }, [])

  useEffect(() => {
    void window.adoApi.getMsWebCounts().then(setWebCounts)
    return window.adoApi.onMsWebCounts(setWebCounts)
  }, [])

  useEffect(() => {
    if (!notice) return
    const t = window.setTimeout(() => setNotice(null), 3600)
    return () => window.clearTimeout(t)
  }, [notice])

  useEffect(() => {
    if (!settings.autoSyncMinutes) return
    const ms = settings.autoSyncMinutes * 60 * 1000
    const id = window.setInterval(() => {
      void syncMine({ quiet: true })
    }, ms)
    return () => window.clearInterval(id)
  }, [settings.autoSyncMinutes, syncMine])

  const updateDraftConnection = (
    connectionId: string,
    patch: Partial<AdoOrgConnection>
  ) => {
    setDraftSettings((s) => ({
      ...s,
      connections: s.connections.map((c) =>
        c.id === connectionId ? { ...c, ...patch } : c
      )
    }))
  }

  const activeConnectionIds = useMemo(() => {
    return new Set(
      settings.connections
        .filter((c) => c.active !== false && c.organization && c.pat)
        .map((c) => c.id)
    )
  }, [settings.connections])

  const sortedMaster = useMemo(() => {
    const q = search.trim().toLowerCase()
    return [...master]
      .filter((e) => {
        if (e.connectionId && !activeConnectionIds.has(e.connectionId)) return false
        if (!q) return true
        return (
          e.title.toLowerCase().includes(q) ||
          e.project.toLowerCase().includes(q) ||
          e.organization.toLowerCase().includes(q) ||
          e.workItemType.toLowerCase().includes(q) ||
          e.state.toLowerCase().includes(q) ||
          String(e.workItemId).includes(q)
        )
      })
      .sort((a, b) => {
        const orgCmp = (a.organization || '').localeCompare(b.organization || '')
        if (orgCmp !== 0) return orgCmp
        const projectCmp = a.project.localeCompare(b.project)
        if (projectCmp !== 0) return projectCmp
        if (a.priority !== b.priority) return a.priority - b.priority
        return a.title.localeCompare(b.title)
      })
  }, [master, search, activeConnectionIds])

  const groupedMaster = useMemo(() => {
    const map = new Map<string, { label: string; entries: MasterEntry[] }>()
    for (const c of settings.connections) {
      if (c.active === false || !c.organization) continue
      const key = c.organization
      map.set(key, {
        label: c.label?.trim() || c.organization,
        entries: []
      })
    }
    for (const entry of sortedMaster) {
      const key = entry.organization || 'Unknown org'
      const group = map.get(key)
      if (group) group.entries.push(entry)
      else map.set(key, { label: key, entries: [entry] })
    }
    return [...map.entries()]
      .sort((a, b) => a[1].label.localeCompare(b[1].label))
      .map(([key, group]) => ({
        key,
        label: group.label,
        entries: group.entries
      }))
  }, [sortedMaster, settings.connections])

  const selected = useMemo(
    () => sortedMaster.find((e) => e.id === selectedId) ?? sortedMaster[0] ?? null,
    [sortedMaster, selectedId]
  )

  useEffect(() => {
    if (!selected?.connectionId || !selected.workItemId) {
      setComments([])
      setCommentsError(null)
      return
    }
    let cancelled = false
    setCommentsLoading(true)
    setCommentsError(null)
    void window.adoApi
      .getWorkItemComments({
        connectionId: selected.connectionId,
        id: selected.workItemId,
        project: selected.project
      })
      .then((list) => {
        if (!cancelled) setComments(list)
      })
      .catch((e) => {
        if (!cancelled) {
          setComments([])
          setCommentsError(
            e instanceof Error ? e.message : 'Failed to load discussion'
          )
        }
      })
      .finally(() => {
        if (!cancelled) setCommentsLoading(false)
      })
    return () => {
      cancelled = true
    }
  }, [selected?.id, selected?.connectionId, selected?.workItemId, selected?.project])

  useEffect(() => {
    if (!selected?.organization) return
    setExpandedOrgs((prev) => {
      if (prev.has(selected.organization)) return prev
      const next = new Set(prev)
      next.add(selected.organization)
      return next
    })
  }, [selected?.id, selected?.organization])

  const toggleOrg = (org: string) => {
    setExpandedOrgs((prev) => {
      const next = new Set(prev)
      if (next.has(org)) next.delete(org)
      else next.add(org)
      return next
    })
  }

  const pinNow = (pin: NowPin) => {
    const result = toggleNowPin(nowPins, pin)
    if (result.error) {
      setNotice(result.error)
      return
    }
    setNowPins(result.pins)
    void window.adoApi.saveNowPins(result.pins)
  }

  const stats = useMemo(() => {
    const orgs = new Set(sortedMaster.map((e) => e.organization).filter(Boolean))
    const projects = new Set(sortedMaster.map((e) => e.project).filter(Boolean))
    return {
      total: sortedMaster.length,
      orgs: orgs.size,
      projects: projects.size
    }
  }, [sortedMaster])

  const saveSettings = async () => {
    setError(null)
    try {
      const saved = await window.adoApi.saveSettings(draftSettings)
      const normalized = normalizeAppSettings(saved)
      setSettings(normalized)
      setDraftSettings(normalized)
      setModal(null)
      if (
        screen !== 'work' &&
        screen !== 'calendar' &&
        screen !== 'teams' &&
        screen !== 'notes' &&
        screen !== 'now' &&
        !normalized.customTabs.some((t) => t.id === screen)
      ) {
        setScreen('work')
      }
      setNotice(
        `Saved ${normalized.connections.length} organization${normalized.connections.length === 1 ? '' : 's'}`
      )
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to save settings')
    }
  }

  const openSettings = async () => {
    const draft = normalizeAppSettings(settings)
    if (draft.connections.length === 0) {
      draft.connections = [emptyConnection()]
    }
    setDraftSettings(draft)
    setModal('settings')
  }

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark">
            ADO <span>Master</span>
          </div>
          <div className="brand-sub">
            {stats.orgs} org{stats.orgs === 1 ? '' : 's'} · {stats.projects} projects
            {settings.autoSyncMinutes > 0 ? ` · auto ${settings.autoSyncMinutes}m` : ''}
            {lastSyncAt ? ` · synced ${formatDate(lastSyncAt)}` : ''}
          </div>
        </div>

        <div className="top-actions">
          {error && (
            <div className="toast toast-error" title={error}>
              <span>{error}</span>
              <button type="button" onClick={() => setError(null)} aria-label="Dismiss">
                ×
              </button>
            </div>
          )}
          {notice && !error && <div className="toast toast-ok">{notice}</div>}
          <button
            className={screen === 'now' ? 'btn' : 'btn btn-ghost'}
            type="button"
            onClick={() => setScreen('now')}
          >
            Now
            {nowPins.length > 0 && <span className="tab-count">{nowPins.length}</span>}
          </button>
          <button
            className={screen === 'work' ? 'btn' : 'btn btn-ghost'}
            type="button"
            onClick={() => setScreen('work')}
          >
            Work
            <span className="tab-count">{stats.total}</span>
          </button>
          <button
            className={screen === 'notes' ? 'btn' : 'btn btn-ghost'}
            type="button"
            onClick={() => setScreen('notes')}
          >
            Notes
            {notesOpenCount > 0 && <span className="tab-count">{notesOpenCount}</span>}
          </button>
          <button
            className={screen === 'calendar' ? 'btn' : 'btn btn-ghost'}
            type="button"
            onClick={() => setScreen('calendar')}
          >
            Outlook
            {webCounts.outlookUnread > 0 && (
              <span className="tab-count tab-count-hot">{webCounts.outlookUnread}</span>
            )}
          </button>
          <button
            className={screen === 'teams' ? 'btn' : 'btn btn-ghost'}
            type="button"
            onClick={() => setScreen('teams')}
          >
            Teams
            {webCounts.teamsUnread > 0 && (
              <span className="tab-count tab-count-hot">{webCounts.teamsUnread}</span>
            )}
          </button>
          <button
            className={alertsMuted ? 'btn btn-mute-on' : 'btn btn-ghost'}
            type="button"
            title={
              alertsMuted
                ? 'Alerts are off. Click to turn ticket, Outlook, and Teams sounds back on.'
                : 'Silence ticket, Outlook, and Teams alerts while you are on a call.'
            }
            onClick={() => void window.adoApi.setAlertsMuted(!alertsMuted)}
          >
            {alertsMuted ? 'Muted' : 'Mute'}
          </button>
          {settings.customTabs.map((tab) => (
            <button
              key={tab.id}
              className={screen === tab.id ? 'btn' : 'btn btn-ghost'}
              type="button"
              onClick={() => setScreen(tab.id)}
            >
              {tab.label}
              {(webCounts.extra?.[tab.id] ?? 0) > 0 && (
                <span className="tab-count tab-count-hot">
                  {webCounts.extra[tab.id]}
                </span>
              )}
            </button>
          ))}
          <button className="btn btn-ghost" onClick={() => void openSettings()}>
            Settings
          </button>
          <button
            className="btn btn-primary"
            onClick={() => void syncMine()}
            disabled={loading || !hasConnection}
          >
            {loading ? 'Working…' : 'Sync'}
          </button>
        </div>
      </header>

      {screen === 'now' ? (
        <NowPane
          master={master}
          pins={nowPins}
          refreshToken={notesEpoch}
          onUnpin={(pin) => pinNow(pin)}
          onOpenWork={(id) => {
            setSelectedId(id)
            setScreen('work')
          }}
          onOpenNote={(id) => {
            setFocusNoteId(id)
            setScreen('notes')
          }}
          onNoteDone={(id, done) => {
            void (async () => {
              const list = await window.adoApi.getLocalNotes()
              const next = list.map((n) =>
                n.id === id
                  ? { ...n, done, updatedAt: new Date().toISOString() }
                  : n
              )
              await window.adoApi.saveLocalNotes(next)
              setNotesOpenCount(next.filter((n) => !n.done).length)
              setNotesEpoch((n) => n + 1)
            })()
          }}
        />
      ) : screen === 'calendar' ? (
        <MsWebEmbed id="outlook" title="Outlook" hidden={modal !== null} />
      ) : screen === 'teams' ? (
        <MsWebEmbed id="teams" title="Teams" hidden={modal !== null} />
      ) : screen === 'notes' ? (
        <NotesPane
          onOpenCount={setNotesOpenCount}
          focusId={focusNoteId}
          pins={nowPins}
          onTogglePin={(id) => pinNow({ kind: 'note', id })}
        />
      ) : settings.customTabs.some((t) => t.id === screen) ? (
        <MsWebEmbed
          id={screen}
          title={settings.customTabs.find((t) => t.id === screen)?.label || 'Tab'}
          url={settings.customTabs.find((t) => t.id === screen)?.url}
          hidden={modal !== null}
        />
      ) : (
      <div className="layout">
        <section className="list-pane">
          <div className="pane-label-row">
            <div className="pane-label">Assigned work</div>
            <input
              className="search"
              placeholder="Filter…"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
          {outlook.connected && (settings.outlookCalendar || settings.outlookMailAlerts) && (
            <div className="outlook-panel">
              {outlook.error && (
                <p className="sidebar-empty">{outlook.error}</p>
              )}
              {settings.outlookCalendar && (
                <>
                  <div className="pane-label">Meetings (Outlook web)</div>
                  {outlook.events.length === 0 ? (
                    <p className="sidebar-empty">No meetings in the next day.</p>
                  ) : (
                    outlook.events.slice(0, 8).map((ev) => (
                      <button
                        key={`${ev.id}-${ev.start}`}
                        type="button"
                        className="meeting-row"
                        onClick={() => void window.adoApi.openExternal(ev.webLink)}
                      >
                        <div className="meeting-when">
                          {formatMeetingWhen(ev.start, ev.end, ev.isAllDay)}
                        </div>
                        <div className="item-row-title">{ev.subject}</div>
                        {ev.location ? (
                          <div className="item-row-meta">{ev.location}</div>
                        ) : null}
                      </button>
                    ))
                  )}
                </>
              )}
              {settings.outlookMailAlerts && (
                <>
                  <div className="pane-label" style={{ marginTop: '0.55rem' }}>
                    Unread mail
                  </div>
                  {outlook.mail.length === 0 ? (
                    <p className="sidebar-empty">Inbox is clear.</p>
                  ) : (
                    outlook.mail.slice(0, 6).map((msg) => (
                      <button
                        key={msg.id}
                        type="button"
                        className="meeting-row"
                        onClick={() => void window.adoApi.openExternal(msg.webLink)}
                      >
                        <div className="item-row-title">{msg.subject}</div>
                        <div className="item-row-meta">{msg.from}</div>
                      </button>
                    ))
                  )}
                </>
              )}
            </div>
          )}
          <div className="item-list">
            {groupedMaster.length === 0 ? (
              <p className="sidebar-empty">
                {hasConnection
                  ? 'Sync to load items assigned to you.'
                  : 'Add an organization in Settings, then Sync.'}
              </p>
            ) : (
              groupedMaster.map((group) => {
                const open =
                  expandedOrgs.has(group.key) ||
                  (Boolean(search.trim()) && group.entries.length > 0)
                return (
                  <div key={group.key} className="project-group">
                    <button
                      type="button"
                      className={`project-group-head${open ? ' open' : ''}`}
                      onClick={() => toggleOrg(group.key)}
                      aria-expanded={open}
                    >
                      <span className="project-group-head-label">
                        <span className="project-group-arrow">▸</span>
                        <span className="project-group-name">{group.label}</span>
                      </span>
                      <span className="project-group-count">{group.entries.length}</span>
                    </button>
                    {open && (
                      <div className="project-group-tickets">
                        {group.entries.length === 0 ? (
                          <p className="sidebar-empty">No tickets assigned.</p>
                        ) : (
                          group.entries.map((entry) => (
                          <button
                            key={entry.id}
                            type="button"
                            className={`item-row ${ticketAgeClass(entry)}${selected?.id === entry.id ? ' selected' : ''}`.trim()}
                            onClick={() => setSelectedId(entry.id)}
                          >
                            <div className="item-row-top">
                              <span className="priority">P{entry.priority}</span>
                              <span className="item-id">#{entry.workItemId}</span>
                              <span className={statusClass(entry.status)}>{entry.status}</span>
                              <span className="chip">{entry.state}</span>
                            </div>
                            <div className="item-row-title">{entry.title}</div>
                            <div className="item-row-meta">
                              {entry.project} · {entry.workItemType}
                              {nowPins.some(
                                (p) => nowPinKey(p) === nowPinKey({ kind: 'work', id: entry.id })
                              )
                                ? ' · Now'
                                : ''}
                            </div>
                          </button>
                        ))
                        )}
                      </div>
                    )}
                  </div>
                )
              })
            )}
          </div>
        </section>

        <main className="detail-pane">
          {!selected ? (
            <div className="detail-empty">Select an item to see full details.</div>
          ) : (
            <>
              <div className="detail-header">
                <div>
                  <div className="detail-kicker">
                    #{selected.workItemId} · {selected.workItemType} ·{' '}
                    {selected.organization} / {selected.project}
                  </div>
                  <h1>{selected.title}</h1>
                </div>
                <div className="row-actions">
                  <button
                    className="btn"
                    type="button"
                    onClick={() => pinNow({ kind: 'work', id: selected.id })}
                  >
                    {nowPins.some((p) => nowPinKey(p) === nowPinKey({ kind: 'work', id: selected.id }))
                      ? 'Unpin from Now'
                      : nowPins.length >= NOW_PIN_LIMIT
                        ? 'Now is full'
                        : 'Pin to Now'}
                  </button>
                  <button
                    className="btn"
                    onClick={() => void window.adoApi.openExternal(selected.webUrl)}
                  >
                    Open
                  </button>
                </div>
              </div>

              <div className="detail-chips">
                <span className={statusClass(selected.status)}>{selected.status}</span>
                <span className="chip">ADO: {selected.state}</span>
                <span className="chip">P{selected.priority}</span>
                <span className="chip">{selected.organization}</span>
                <span className="chip">{selected.project}</span>
              </div>

              <dl className="detail-meta people-meta">
                <div>
                  <dt>Assigned to</dt>
                  <dd>{selected.assignedTo || 'Unassigned'}</dd>
                </div>
                <div>
                  <dt>Created by</dt>
                  <dd>{selected.createdBy || '—'}</dd>
                </div>
                <div>
                  <dt>Created</dt>
                  <dd>{formatDate(selected.createdDate)}</dd>
                </div>
                <div>
                  <dt>Last changed by</dt>
                  <dd>{selected.changedBy || '—'}</dd>
                </div>
              </dl>

              <div className="detail-stack">
              <div className="detail-block">
                <h3>Description</h3>
                <RichHtml
                  className="detail-body"
                  html={selected.description}
                  organization={selected.organization}
                  empty="No description."
                />
              </div>

              <div className="detail-block discussion-block">
                <h3>Discussion {comments.length > 0 ? `(${comments.length})` : ''}</h3>
                {commentsLoading && (
                  <div className="detail-body">Loading comments…</div>
                )}
                {commentsError && (
                  <div className="detail-body" style={{ color: 'var(--danger)' }}>
                    {commentsError}
                  </div>
                )}
                {!commentsLoading && !commentsError && comments.length === 0 && (
                  <div className="detail-body">No discussion yet.</div>
                )}
                {!commentsLoading && comments.length > 0 && (
                  <div className="discussion-list">
                    {comments.map((c) => (
                      <article key={c.id} className="discussion-item">
                        <header>
                          <strong>{c.createdBy || 'Unknown'}</strong>
                          <span>{formatDate(c.createdDate)}</span>
                        </header>
                        <RichHtml
                          html={c.text}
                          organization={selected.organization}
                          empty="(empty comment)"
                        />
                      </article>
                    ))}
                  </div>
                )}
              </div>
              </div>

              <dl className="detail-meta">
                <div>
                  <dt>Changed</dt>
                  <dd>{formatDate(selected.changedDate)}</dd>
                </div>
                <div>
                  <dt>Added to list</dt>
                  <dd>{formatDate(selected.addedAt)}</dd>
                </div>
                <div>
                  <dt>Organization</dt>
                  <dd>{selected.organization}</dd>
                </div>
                <div>
                  <dt>Project</dt>
                  <dd>{selected.project}</dd>
                </div>
              </dl>
            </>
          )}
        </main>
      </div>
      )}

      {modal === 'settings' && (
        <div className="modal-backdrop" onClick={() => setModal(null)}>
          <div className="modal modal-wide modal-setup" onClick={(e) => e.stopPropagation()}>
            <h2>Settings</h2>
            <div className="field">
              <label>Appearance</label>
              <div className="setup-sound-row">
                <button
                  className={theme === 'light' ? 'btn' : 'btn btn-ghost'}
                  type="button"
                  onClick={() => setTheme('light')}
                >
                  Light
                </button>
                <button
                  className={theme === 'dark' ? 'btn' : 'btn btn-ghost'}
                  type="button"
                  onClick={() => setTheme('dark')}
                >
                  Dark
                </button>
              </div>
            </div>
            <p className="hint">
              One row per Azure DevOps organization. Check Active to include it in Sync.
              Every project in an active org is included.
            </p>

            <div className="setup-grid">
              <div className="setup-grid-head">
                <span>Organization</span>
                <span>PAT</span>
                <span>Label</span>
                <span>Active</span>
                <span />
              </div>
              {draftSettings.connections.map((conn) => (
                <div key={conn.id} className="setup-row-wrap">
                  <div className="setup-row">
                    <input
                      placeholder="https://dev.azure.com/my-org"
                      value={conn.organization}
                      onChange={(e) =>
                        updateDraftConnection(conn.id, {
                          organization: e.target.value
                        })
                      }
                    />
                    <input
                      type="password"
                      placeholder="Personal access token"
                      value={conn.pat}
                      onChange={(e) =>
                        updateDraftConnection(conn.id, { pat: e.target.value })
                      }
                    />
                    <input
                      placeholder="Optional"
                      value={conn.label ?? ''}
                      onChange={(e) =>
                        updateDraftConnection(conn.id, { label: e.target.value })
                      }
                    />
                    <label className="setup-active">
                      <input
                        type="checkbox"
                        checked={conn.active !== false}
                        onChange={(e) =>
                          updateDraftConnection(conn.id, { active: e.target.checked })
                        }
                      />
                      <span>{conn.active !== false ? 'On' : 'Off'}</span>
                    </label>
                    <button
                      className="btn btn-danger btn-tiny"
                      type="button"
                      aria-label="Remove organization"
                      onClick={() =>
                        setDraftSettings((s) => ({
                          ...s,
                          connections: s.connections.filter((c) => c.id !== conn.id)
                        }))
                      }
                    >
                      −
                    </button>
                  </div>
                </div>
              ))}
            </div>

            <div className="row-actions" style={{ marginBottom: '0.9rem' }}>
              <button
                className="btn btn-primary"
                type="button"
                onClick={() => {
                  const conn = emptyConnection()
                  setDraftSettings((s) => ({
                    ...s,
                    connections: [...s.connections, conn]
                  }))
                }}
              >
                Add row
              </button>
            </div>

            <div className="field">
              <label htmlFor="auto-sync">Auto-sync interval</label>
              <select
                id="auto-sync"
                value={draftSettings.autoSyncMinutes}
                onChange={(e) =>
                  setDraftSettings((s) => ({
                    ...s,
                    autoSyncMinutes: Number(e.target.value)
                  }))
                }
              >
                <option value={0}>Off</option>
                <option value={2}>Every 2 minutes</option>
                <option value={5}>Every 5 minutes</option>
                <option value={15}>Every 15 minutes</option>
                <option value={30}>Every 30 minutes</option>
                <option value={60}>Every 60 minutes</option>
              </select>
              <p className="hint">
                Syncs assigned work items and refreshes Outlook and Teams unread counts.
              </p>
            </div>

            <div className="field">
              <label htmlFor="alert-sound">New ticket sound</label>
              <div className="setup-sound-row">
                <select
                  id="alert-sound"
                  value={draftSettings.alertSound || 'Alarm'}
                  onChange={(e) =>
                    setDraftSettings((s) => ({
                      ...s,
                      alertSound: e.target.value
                    }))
                  }
                >
                  {ALERT_SOUNDS.map((name) => (
                    <option key={name} value={name}>
                      {ALERT_SOUND_LABELS[name] ?? name}
                    </option>
                  ))}
                </select>
                <button
                  className="btn"
                  type="button"
                  disabled={(draftSettings.alertSound || 'Alarm') === 'none'}
                  onClick={() =>
                    void window.adoApi.playAlertSound(draftSettings.alertSound)
                  }
                >
                  Test
                </button>
              </div>
            </div>

            <div className="field">
              <label>Outlook and Teams</label>
              <p className="hint">
                Outlook and Teams load Microsoft’s web apps inside this window.
                Sign in with your Microsoft 365 account there — no app
                registration. The session stays on this Mac and is shared
                between both tabs.
              </p>
              <label className="setup-active" style={{ justifyContent: 'flex-start', marginTop: '0.55rem' }}>
                <input
                  type="checkbox"
                  checked={draftSettings.webActivityAlerts !== false}
                  onChange={(e) =>
                    setDraftSettings((s) => ({
                      ...s,
                      webActivityAlerts: e.target.checked
                    }))
                  }
                />
                <span>Loud alert for new Outlook mail and Teams messages</span>
              </label>
              <div style={{ marginTop: '0.65rem' }}>
                <AlertSoundPicker
                  id="outlook-alert-sound"
                  label="Outlook mail sound"
                  value={draftSettings.outlookAlertSound}
                  onChange={(outlookAlertSound) =>
                    setDraftSettings((s) => ({ ...s, outlookAlertSound }))
                  }
                />
                <AlertSoundPicker
                  id="teams-alert-sound"
                  label="Teams message sound"
                  value={draftSettings.teamsAlertSound}
                  onChange={(teamsAlertSound) =>
                    setDraftSettings((s) => ({ ...s, teamsAlertSound }))
                  }
                />
              </div>
              <p className="hint" style={{ marginTop: '0.15rem' }}>
                ADHD options (red alert, air horn, ringing phone, strobe, two-tone,
                and spoken mail/Teams lines) play extra loud. Turn on notifications
                inside Outlook and Teams (web). Open each tab once so both stay
                signed in in the background. Save after you pick sounds.
              </p>
            </div>

            <div className="field">
              <label>Custom tabs</label>
              <p className="hint">
                Add extra tabs that load any https URL in the app, like Outlook
                and Teams.
              </p>
              <div className="setup-grid">
                <div className="setup-grid-head tab-setup-head">
                  <span>Name</span>
                  <span>URL</span>
                  <span />
                </div>
                {draftSettings.customTabs.map((tab) => (
                  <div key={tab.id} className="setup-row-wrap">
                    <div className="setup-row tab-setup-row">
                      <input
                        placeholder="Name"
                        value={tab.label}
                        onChange={(e) =>
                          setDraftSettings((s) => ({
                            ...s,
                            customTabs: s.customTabs.map((t) =>
                              t.id === tab.id ? { ...t, label: e.target.value } : t
                            )
                          }))
                        }
                      />
                      <input
                        placeholder="https://example.com"
                        value={tab.url}
                        onChange={(e) =>
                          setDraftSettings((s) => ({
                            ...s,
                            customTabs: s.customTabs.map((t) =>
                              t.id === tab.id ? { ...t, url: e.target.value } : t
                            )
                          }))
                        }
                      />
                      <button
                        className="btn btn-danger btn-tiny"
                        type="button"
                        aria-label="Remove tab"
                        onClick={() =>
                          setDraftSettings((s) => ({
                            ...s,
                            customTabs: s.customTabs.filter((t) => t.id !== tab.id)
                          }))
                        }
                      >
                        −
                      </button>
                    </div>
                  </div>
                ))}
              </div>
              <button
                className="btn"
                type="button"
                onClick={() =>
                  setDraftSettings((s) => ({
                    ...s,
                    customTabs: [...s.customTabs, emptyCustomTab()]
                  }))
                }
              >
                Add tab
              </button>
            </div>

            <div className="modal-actions">
              <button className="btn btn-ghost" onClick={() => setModal(null)}>
                Cancel
              </button>
              <button className="btn btn-primary" onClick={() => void saveSettings()}>
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
