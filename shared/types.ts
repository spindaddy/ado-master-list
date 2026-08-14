export interface AdoOrgConnection {
  id: string
  /** Org slug, e.g. contoso */
  organization: string
  pat: string
  /**
   * Project names included for this org.
   * Unused: sync always includes all projects.
   */
  enabledProjects: string[]
  /** When false, this org is skipped on sync and hidden from the list. */
  active: boolean
  /** Optional display label */
  label?: string
}

export interface AppSettings {
  connections: AdoOrgConnection[]
  /** 0 = off. Otherwise minutes between automatic syncs. */
  autoSyncMinutes: number
  /** macOS system sound name, or "none". */
  alertSound: string
  /** Entra / Azure AD application (client) ID for Microsoft Graph. */
  outlookClientId: string
  outlookCalendar: boolean
  outlookMailAlerts: boolean
  /** Minutes before a meeting to fire the alert sound. 0 = at start. */
  outlookMeetingAlertMinutes: number
  /** Loud alert when Outlook or Teams web shows a notification. */
  webActivityAlerts: boolean
  outlookAlertSound: string
  teamsAlertSound: string
  customTabs: CustomWebTab[]
}

export interface CustomWebTab {
  id: string
  label: string
  url: string
}

export function normalizeWebTabUrl(raw: string): string {
  const value = raw.trim()
  if (!value) return ''
  if (/^(javascript|data|file|about|blob):/i.test(value)) return ''
  if (/^https?:\/\//i.test(value)) return value
  return `https://${value}`
}

export function normalizeCustomWebTab(
  raw: Partial<CustomWebTab> | null | undefined
): CustomWebTab | null {
  if (!raw) return null
  const url = normalizeWebTabUrl(String(raw.url ?? ''))
  if (!url) return null
  try {
    const parsed = new URL(url)
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return null
    const id = String(raw.id ?? '').trim() || `tab-${Date.now().toString(36)}`
    const label = String(raw.label ?? '').trim() || parsed.hostname.replace(/^www\./, '')
    return { id, label, url }
  } catch {
    return null
  }
}

export interface OutlookEventDto {
  id: string
  subject: string
  start: string
  end: string
  location: string
  isAllDay: boolean
  webLink: string
  organizer: string
}

export interface OutlookMailDto {
  id: string
  subject: string
  from: string
  receivedDateTime: string
  webLink: string
  isRead: boolean
}

export interface OutlookSnapshot {
  connected: boolean
  account: string
  error: string | null
  events: OutlookEventDto[]
  mail: OutlookMailDto[]
  unreadCount: number
}

export const ALERT_SOUNDS = [
  'none',
  'Alarm',
  'TripleBeep',
  'Klaxon',
  'RedAlert',
  'AirHorn',
  'PhoneNag',
  'StrobeBeep',
  'TwoTone',
  'Voice',
  'VoiceAlarm',
  'VoiceMail',
  'VoiceMailAlarm',
  'VoiceTeams',
  'VoiceTeamsAlarm',
  'Basso',
  'Blow',
  'Bottle',
  'Frog',
  'Funk',
  'Glass',
  'Hero',
  'Morse',
  'Ping',
  'Pop',
  'Purr',
  'Sosumi',
  'Submarine',
  'Tink'
] as const

export const ALERT_SOUND_LABELS: Record<string, string> = {
  none: 'None',
  Alarm: 'Loud alarm (siren)',
  TripleBeep: 'Loud triple beep',
  Klaxon: 'Loud klaxon',
  RedAlert: 'ADHD red alert (long siren)',
  AirHorn: 'ADHD air horn',
  PhoneNag: 'ADHD ringing phone',
  StrobeBeep: 'ADHD strobe beeps',
  TwoTone: 'ADHD two-tone (ambulance)',
  Voice: 'Spoken “New ticket”',
  VoiceAlarm: 'Voice + loud alarm',
  VoiceMail: 'Spoken “New Outlook mail”',
  VoiceMailAlarm: 'Outlook voice + red alert',
  VoiceTeams: 'Spoken “New Teams message”',
  VoiceTeamsAlarm: 'Teams voice + air horn',
  Basso: 'Basso (system)',
  Blow: 'Blow (system)',
  Bottle: 'Bottle (system)',
  Frog: 'Frog (system)',
  Funk: 'Funk (system)',
  Glass: 'Glass (system, quiet)',
  Hero: 'Hero (system)',
  Morse: 'Morse (system)',
  Ping: 'Ping (system)',
  Pop: 'Pop (system)',
  Purr: 'Purr (system)',
  Sosumi: 'Sosumi (system)',
  Submarine: 'Submarine (system)',
  Tink: 'Tink (system, quiet)'
}

export const OUTLOOK_REDIRECT_URI = 'http://localhost:18764/auth'
export const OUTLOOK_GRAPH_SCOPES = 'Calendars.Read, Mail.ReadBasic, User.Read, offline_access'

export function normalizeAlertSound(value: unknown): string {
  const name = String(value ?? 'Alarm')
  return (ALERT_SOUNDS as readonly string[]).includes(name) ? name : 'Alarm'
}

export interface AdoProjectDto {
  id: string
  name: string
  description: string
  state: string
  organization: string
  connectionId: string
}

export type WorkStatus = 'backlog' | 'active' | 'blocked' | 'done'

export interface AdoWorkItemDto {
  id: number
  title: string
  workItemType: string
  state: string
  project: string
  organization: string
  connectionId: string
  assignedTo: string
  createdBy: string
  createdDate?: string
  changedBy: string
  changedDate?: string
  description: string
  webUrl: string
}

export interface AdoCommentDto {
  id: number
  text: string
  createdBy: string
  createdDate: string
  modifiedDate?: string
}

export interface MasterEntry {
  /** Unique across orgs: `${organization}:${workItemId}` */
  id: string
  workItemId: number
  title: string
  workItemType: string
  state: string
  project: string
  organization: string
  connectionId: string
  assignedTo: string
  createdBy: string
  createdDate?: string
  changedBy: string
  description: string
  webUrl: string
  changedDate?: string
  /** Local working fields */
  notes: string
  status: WorkStatus
  priority: number
  addedAt: string
}

export interface MasterListState {
  entries: MasterEntry[]
}
