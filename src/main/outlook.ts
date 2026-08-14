import { BrowserWindow } from 'electron'
import { createHash, randomBytes } from 'crypto'
import { createServer, type IncomingMessage, type ServerResponse } from 'http'
import { OUTLOOK_REDIRECT_URI } from '../../shared/types'
import type { OutlookEventDto, OutlookMailDto, OutlookSnapshot } from '../../shared/types'

export const OUTLOOK_SCOPES =
  'offline_access User.Read Calendars.Read Mail.ReadBasic'

const AUTH_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
const TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
const GRAPH = 'https://graph.microsoft.com/v1.0'
const AUTH_PORT = 18764

export interface OutlookTokens {
  clientId: string
  accessToken: string
  refreshToken: string
  expiresAt: number
  account: string
}

function base64Url(buf: Buffer): string {
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '')
}

function pkce(): { verifier: string; challenge: string } {
  const verifier = base64Url(randomBytes(32))
  const challenge = base64Url(createHash('sha256').update(verifier).digest())
  return { verifier, challenge }
}

async function readRequestUrl(req: IncomingMessage): Promise<URL> {
  const host = req.headers.host || `localhost:${AUTH_PORT}`
  return new URL(req.url || '/', `http://${host}`)
}

function sendHtml(res: ServerResponse, status: number, body: string): void {
  res.writeHead(status, { 'Content-Type': 'text/html; charset=utf-8' })
  res.end(
    `<!doctype html><html><body style="font-family:system-ui;padding:2rem">
     <p>${body}</p></body></html>`
  )
}

interface TokenResponse {
  access_token?: string
  refresh_token?: string
  expires_in?: number
  error?: string
  error_description?: string
}

async function exchangeToken(body: Record<string, string>): Promise<TokenResponse> {
  const res = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(body).toString()
  })
  return (await res.json()) as TokenResponse
}

function tokensFromResponse(
  clientId: string,
  data: TokenResponse,
  previous?: OutlookTokens
): OutlookTokens {
  if (!data.access_token) {
    throw new Error(data.error_description || data.error || 'Microsoft sign-in failed')
  }
  return {
    clientId,
    accessToken: data.access_token,
    refreshToken: data.refresh_token || previous?.refreshToken || '',
    expiresAt: Date.now() + Math.max(60, Number(data.expires_in) || 3600) * 1000 - 60_000,
    account: previous?.account || ''
  }
}

export async function signInToOutlook(
  clientId: string,
  parent: BrowserWindow | null
): Promise<OutlookTokens> {
  const id = clientId.trim()
  if (!id) throw new Error('Paste an Entra app client ID in Settings first.')

  const { verifier, challenge } = pkce()
  const state = base64Url(randomBytes(16))
  const authUrl =
    `${AUTH_URL}?` +
    new URLSearchParams({
      client_id: id,
      response_type: 'code',
      redirect_uri: OUTLOOK_REDIRECT_URI,
      response_mode: 'query',
      scope: OUTLOOK_SCOPES,
      code_challenge: challenge,
      code_challenge_method: 'S256',
      state,
      prompt: 'select_account'
    }).toString()

  const code = await new Promise<string>((resolve, reject) => {
    const server = createServer(async (req, res) => {
      try {
        const url = await readRequestUrl(req)
        if (url.pathname !== '/auth') {
          sendHtml(res, 404, 'Not found')
          return
        }
        const err = url.searchParams.get('error_description') || url.searchParams.get('error')
        if (err) {
          sendHtml(res, 400, err)
          cleanup()
          reject(new Error(err))
          return
        }
        if (url.searchParams.get('state') !== state) {
          sendHtml(res, 400, 'Sign-in state mismatch.')
          cleanup()
          reject(new Error('Sign-in state mismatch.'))
          return
        }
        const nextCode = url.searchParams.get('code')
        if (!nextCode) {
          sendHtml(res, 400, 'No authorization code returned.')
          cleanup()
          reject(new Error('No authorization code returned.'))
          return
        }
        sendHtml(res, 200, 'Signed in. You can close this window and return to ADO Master.')
        cleanup()
        resolve(nextCode)
      } catch (e) {
        cleanup()
        reject(e instanceof Error ? e : new Error('Sign-in failed'))
      }
    })

    const win = new BrowserWindow({
      width: 520,
      height: 740,
      parent: parent ?? undefined,
      modal: Boolean(parent),
      title: 'Sign in to Outlook',
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: true
      }
    })

    let settled = false
    const cleanup = () => {
      if (settled) return
      settled = true
      try {
        server.close()
      } catch {
        // ignore
      }
      if (!win.isDestroyed()) win.close()
    }

    server.once('error', (e) => {
      cleanup()
      reject(e)
    })
    win.on('closed', () => {
      if (!settled) {
        settled = true
        try {
          server.close()
        } catch {
          // ignore
        }
        reject(new Error('Sign-in window closed.'))
      }
    })

    server.listen(AUTH_PORT, () => {
      void win.loadURL(authUrl)
    })

    setTimeout(() => {
      if (settled) return
      cleanup()
      reject(new Error('Sign-in timed out.'))
    }, 5 * 60 * 1000)
  })

  const data = await exchangeToken({
    client_id: id,
    grant_type: 'authorization_code',
    code,
    redirect_uri: OUTLOOK_REDIRECT_URI,
    code_verifier: verifier
  })
  const tokens = tokensFromResponse(id, data)
  const me = await graphGet<{ displayName?: string; mail?: string; userPrincipalName?: string }>(
    tokens.accessToken,
    '/me?$select=displayName,mail,userPrincipalName'
  )
  tokens.account =
    me.mail || me.userPrincipalName || me.displayName || 'Microsoft account'
  return tokens
}

export async function refreshOutlookTokens(previous: OutlookTokens): Promise<OutlookTokens> {
  if (!previous.refreshToken) throw new Error('Outlook session expired. Sign in again.')
  const data = await exchangeToken({
    client_id: previous.clientId,
    grant_type: 'refresh_token',
    refresh_token: previous.refreshToken,
    scope: OUTLOOK_SCOPES
  })
  const next = tokensFromResponse(previous.clientId, data, previous)
  next.account = previous.account
  return next
}

async function graphGet<T>(accessToken: string, path: string): Promise<T> {
  const res = await fetch(`${GRAPH}${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Prefer: 'outlook.timezone="UTC"'
    }
  })
  const text = await res.text()
  if (!res.ok) {
    let detail = text
    try {
      const json = JSON.parse(text) as { error?: { message?: string } }
      detail = json.error?.message || text
    } catch {
      // keep text
    }
    const err = new Error(detail || `Graph ${res.status}`) as Error & { status?: number }
    err.status = res.status
    throw err
  }
  return (text ? JSON.parse(text) : {}) as T
}

function graphDate(value?: { dateTime?: string; date?: string }): string {
  return value?.dateTime || value?.date || ''
}

export async function loadOutlookSnapshot(
  accessToken: string,
  opts: { calendar: boolean; mail: boolean }
): Promise<Omit<OutlookSnapshot, 'connected' | 'account' | 'error'>> {
  const events: OutlookEventDto[] = []
  const mail: OutlookMailDto[] = []

  if (opts.calendar) {
    const start = new Date()
    start.setMinutes(0, 0, 0)
    const end = new Date(start.getTime() + 36 * 60 * 60 * 1000)
    const params = new URLSearchParams({
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      $top: '40',
      $select: 'id,subject,start,end,location,isAllDay,webLink,organizer'
    })
    const data = await graphGet<{
      value?: Array<{
        id?: string
        subject?: string
        start?: { dateTime?: string; date?: string }
        end?: { dateTime?: string; date?: string }
        location?: { displayName?: string }
        isAllDay?: boolean
        webLink?: string
        organizer?: { emailAddress?: { name?: string; address?: string } }
      }>
    }>(accessToken, `/me/calendarView?${params.toString()}`)
    for (const ev of data.value ?? []) {
      if (!ev.id) continue
      events.push({
        id: ev.id,
        subject: ev.subject || '(no subject)',
        start: graphDate(ev.start),
        end: graphDate(ev.end),
        location: ev.location?.displayName || '',
        isAllDay: Boolean(ev.isAllDay),
        webLink: ev.webLink || 'https://outlook.office.com/calendar',
        organizer: ev.organizer?.emailAddress?.name || ev.organizer?.emailAddress?.address || ''
      })
    }
    events.sort((a, b) => a.start.localeCompare(b.start))
  }

  if (opts.mail) {
    const params = new URLSearchParams({
      $top: '25',
      $orderby: 'receivedDateTime desc',
      $select: 'id,subject,from,receivedDateTime,webLink,isRead'
    })
    const data = await graphGet<{
      value?: Array<{
        id?: string
        subject?: string
        from?: { emailAddress?: { name?: string; address?: string } }
        receivedDateTime?: string
        webLink?: string
        isRead?: boolean
      }>
    }>(accessToken, `/me/mailFolders/inbox/messages?${params.toString()}`)
    for (const msg of data.value ?? []) {
      if (!msg.id) continue
      mail.push({
        id: msg.id,
        subject: msg.subject || '(no subject)',
        from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || '',
        receivedDateTime: msg.receivedDateTime || '',
        webLink: msg.webLink || 'https://outlook.office.com/mail',
        isRead: Boolean(msg.isRead)
      })
    }
  }

  return {
    events,
    mail,
    unreadCount: mail.filter((m) => !m.isRead).length
  }
}

export function emptyOutlookSnapshot(error: string | null = null): OutlookSnapshot {
  return {
    connected: false,
    account: '',
    error,
    events: [],
    mail: [],
    unreadCount: 0
  }
}
