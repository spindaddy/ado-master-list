import { BrowserView, BrowserWindow, shell, type WebContents } from 'electron'

const CHROME_UA =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'

export const MS_WEB_EMBEDS = {
  outlook: {
    url: 'https://outlook.office.com/calendar/view/workweek'
  },
  teams: {
    url: 'https://teams.microsoft.com/'
  }
} as const

export type MsWebEmbedId = string

const BUILTIN_URLS: Record<string, string> = {
  outlook: 'https://outlook.office.com/calendar/view/workweek',
  teams: 'https://teams.microsoft.com/'
}

export interface MsWebNotify {
  source: MsWebEmbedId
  title: string
  body: string
}

const PARTITION = 'persist:ms365-web'
const NOTIFY_PREFIX = 'ADO_MASTER_NOTIFY::'

const HOOK_JS = `(() => {
  const Orig = window.Notification;
  if (!Orig || Orig.__adoMasterHook) return;
  const Wrapped = function(title, options) {
    try {
      console.debug('${NOTIFY_PREFIX}' + JSON.stringify({
        title: String(title || ''),
        body: String((options && options.body) || ''),
        tag: String((options && options.tag) || '')
      }));
    } catch (e) {}
    return new Orig(title, options);
  };
  Wrapped.__adoMasterHook = true;
  Wrapped.permission = Orig.permission;
  Wrapped.requestPermission = function() {
    return Orig.requestPermission.apply(Orig, arguments);
  };
  window.Notification = Wrapped;
  try { Orig.requestPermission(); } catch (e) {}
})();`

const views = new Map<MsWebEmbedId, BrowserView>()
let attachedWin: BrowserWindow | null = null
let attachedId: MsWebEmbedId | null = null
let permissionsHooked = false
let notifyHandler: ((payload: MsWebNotify) => void) | null = null

export function setMsWebNotifyHandler(handler: (payload: MsWebNotify) => void): void {
  notifyHandler = handler
}

export interface MsWebCounts {
  outlookUnread: number
  teamsUnread: number
  extra: Record<string, number>
}

const COUNT_JS = `(() => {
  const title = document.title || '';
  const titleMatch = title.match(/\\((\\d+)\\)/);
  let best = titleMatch ? Number(titleMatch[1]) : 0;
  const nodes = document.querySelectorAll('[aria-label], [title]');
  for (const el of nodes) {
    const text = [el.getAttribute('aria-label') || '', el.getAttribute('title') || ''].join(' ');
    const m =
      text.match(/(\\d+)\\s+unread/i) ||
      text.match(/unread[^0-9]{0,16}(\\d+)/i) ||
      text.match(/(\\d+)\\s+new (mail|message|messages|chats?)/i);
    if (m) best = Math.max(best, Number(m[1]));
  }
  return Number.isFinite(best) ? best : 0;
})()`

let counts: MsWebCounts = { outlookUnread: 0, teamsUnread: 0, extra: {} }
let countHandler: ((next: MsWebCounts) => void) | null = null
let countTimer: ReturnType<typeof setInterval> | null = null

export function setMsWebCountHandler(handler: (next: MsWebCounts) => void): void {
  countHandler = handler
}

export function getMsWebCounts(): MsWebCounts {
  return counts
}

async function readUnread(id: MsWebEmbedId): Promise<number> {
  const view = views.get(id)
  if (!view || view.webContents.isDestroyed()) return 0
  try {
    const value = await view.webContents.executeJavaScript(COUNT_JS, true)
    const n = Number(value)
    return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0
  } catch {
    return 0
  }
}

async function refreshCounts(): Promise<void> {
  const extra: Record<string, number> = {}
  for (const id of views.keys()) {
    if (id === 'outlook' || id === 'teams') continue
    extra[id] = await readUnread(id)
  }
  const next: MsWebCounts = {
    outlookUnread: await readUnread('outlook'),
    teamsUnread: await readUnread('teams'),
    extra
  }
  if (JSON.stringify(next) === JSON.stringify(counts)) return
  counts = next
  countHandler?.(counts)
}

export function startMsWebCountPolling(): void {
  if (countTimer) return
  countTimer = setInterval(() => {
    void refreshCounts()
  }, 8000)
}

/** Reload Outlook/Teams in the background and re-read unread badges. */
export async function syncMsWebCounts(): Promise<MsWebCounts> {
  for (const id of ['outlook', 'teams'] as const) {
    const existed = views.has(id)
    const view = ensureView(id)
    if (attachedId === id) continue
    if (!existed) continue
    if (view.webContents.isDestroyed() || view.webContents.isLoading()) continue
    const url = view.webContents.getURL() || urlFor(id)
    if (!url) continue
    try {
      await view.webContents.loadURL(url)
    } catch {
      // ignore navigation failures
    }
  }
  await refreshCounts()
  return counts
}

function hookPermissions(sess: WebContents['session']): void {
  if (permissionsHooked) return
  permissionsHooked = true
  sess.setPermissionRequestHandler((_wc, permission, callback) => {
    callback(
      permission === 'media' ||
        permission === 'display-capture' ||
        permission === 'notifications' ||
        permission === 'clipboard-sanitized-write'
    )
  })
}

function injectNotifyHook(wc: WebContents): void {
  const run = (code: string) => {
    try {
      const frames = wc.mainFrame?.framesInSubtree
      if (frames && frames.length > 0) {
        for (const frame of frames) {
          void frame.executeJavaScript(code, true).catch(() => undefined)
        }
        return
      }
    } catch {
      // fall through
    }
    void wc.executeJavaScript(code, true).catch(() => undefined)
  }
  run(HOOK_JS)
}

function parseNotify(message: string): { title: string; body: string } | null {
  const idx = message.indexOf(NOTIFY_PREFIX)
  if (idx < 0) return null
  try {
    const data = JSON.parse(message.slice(idx + NOTIFY_PREFIX.length)) as {
      title?: string
      body?: string
    }
    return { title: data.title || '', body: data.body || '' }
  } catch {
    return null
  }
}

function applyTitleCount(id: MsWebEmbedId, title: string): void {
  const match = String(title || '').match(/\((\d+)\)/)
  if (!match) return
  const n = Number(match[1])
  if (!Number.isFinite(n)) return
  const next: MsWebCounts = {
    outlookUnread: id === 'outlook' ? n : counts.outlookUnread,
    teamsUnread: id === 'teams' ? n : counts.teamsUnread,
    extra: { ...counts.extra }
  }
  if (id !== 'outlook' && id !== 'teams') next.extra[id] = n
  if (JSON.stringify(next) === JSON.stringify(counts)) return
  counts = next
  countHandler?.(counts)
}

function wireView(id: MsWebEmbedId, view: BrowserView): void {
  hookPermissions(view.webContents.session)
  view.webContents.setUserAgent(CHROME_UA)
  view.webContents.on('did-finish-load', () => {
    injectNotifyHook(view.webContents)
    void refreshCounts()
  })
  view.webContents.on('did-frame-finish-load', () => injectNotifyHook(view.webContents))
  view.webContents.on('page-title-updated', (_e, title) => {
    applyTitleCount(id, title)
  })
  view.webContents.on('console-message', (_e, _level, message) => {
    const parsed = parseNotify(String(message || ''))
    if (!parsed) return
    notifyHandler?.({ source: id, title: parsed.title, body: parsed.body })
    void refreshCounts()
  })
  view.webContents.on('did-create-window', (child) => {
    injectNotifyHook(child.webContents)
    child.webContents.on('did-finish-load', () => injectNotifyHook(child.webContents))
    child.webContents.on('console-message', (_e, _level, message) => {
      const parsed = parseNotify(String(message || ''))
      if (!parsed) return
      notifyHandler?.({ source: id, title: parsed.title, body: parsed.body })
      void refreshCounts()
    })
  })
  view.webContents.setWindowOpenHandler(({ url }) => {
    try {
      const parsed = new URL(url)
      if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
        return { action: 'deny' }
      }
    } catch {
      void shell.openExternal(url)
      return { action: 'deny' }
    }
    return {
      action: 'allow',
      overrideBrowserWindowOptions: {
        width: 1100,
        height: 760,
        webPreferences: {
          partition: id === 'outlook' || id === 'teams' ? PARTITION : `persist:webtab-${id}`,
          nodeIntegration: false,
          contextIsolation: true,
          sandbox: false,
          backgroundThrottling: false
        }
      }
    }
  })
}

const viewUrls = new Map<string, string>()

function urlFor(id: MsWebEmbedId, url?: string): string {
  const next = (url || BUILTIN_URLS[id] || viewUrls.get(id) || '').trim()
  return next
}

function ensureView(id: MsWebEmbedId, url?: string): BrowserView {
  const target = urlFor(id, url)
  const existing = views.get(id)
  if (existing) {
    if (target && viewUrls.get(id) !== target) {
      viewUrls.set(id, target)
      void existing.webContents.loadURL(target)
    }
    return existing
  }
  const view = new BrowserView({
    webPreferences: {
      partition: id === 'outlook' || id === 'teams' ? PARTITION : `persist:webtab-${id}`,
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
      backgroundThrottling: false
    }
  })
  view.setBackgroundColor('#ffffff')
  wireView(id, view)
  if (target) {
    viewUrls.set(id, target)
    void view.webContents.loadURL(target)
  }
  views.set(id, view)
  return view
}

function warmOther(id: MsWebEmbedId): void {
  if (id !== 'outlook' && id !== 'teams') return
  const other: MsWebEmbedId = id === 'teams' ? 'outlook' : 'teams'
  setTimeout(() => {
    ensureView(other)
  }, 2000)
}

export function showMsWebEmbed(
  win: BrowserWindow,
  id: MsWebEmbedId,
  bounds: { x: number; y: number; width: number; height: number },
  url?: string
): void {
  const next = ensureView(id, url)
  warmOther(id)
  if (attachedWin && attachedWin !== win && !attachedWin.isDestroyed()) {
    attachedWin.setBrowserView(null)
  }
  if (win.getBrowserView() !== next) {
    win.setBrowserView(next)
  }
  attachedWin = win
  attachedId = id
  next.setBounds({
    x: Math.max(0, Math.round(bounds.x)),
    y: Math.max(0, Math.round(bounds.y)),
    width: Math.max(1, Math.round(bounds.width)),
    height: Math.max(1, Math.round(bounds.height))
  })
  next.webContents.focus()
}

export function hideMsWebEmbed(win: BrowserWindow | null): void {
  if (!win || win.isDestroyed()) return
  win.setBrowserView(null)
  if (attachedWin === win) attachedWin = null
}

export function focusMsWebEmbed(): void {
  if (attachedId) views.get(attachedId)?.webContents.focus()
}

export function reloadMsWebEmbed(id: MsWebEmbedId, url?: string): void {
  const target = urlFor(id, url)
  if (!target) return
  viewUrls.set(id, target)
  void ensureView(id, target).webContents.loadURL(target)
}

export function pruneMsWebEmbeds(keepIds: string[]): void {
  const keep = new Set(['outlook', 'teams', ...keepIds])
  for (const id of [...views.keys()]) {
    if (keep.has(id)) continue
    const view = views.get(id)
    if (view && attachedId === id && attachedWin && !attachedWin.isDestroyed()) {
      attachedWin.setBrowserView(null)
      attachedId = null
    }
    views.delete(id)
    viewUrls.delete(id)
  }
}

export function disposeMsWebEmbeds(): void {
  if (attachedWin && !attachedWin.isDestroyed()) attachedWin.setBrowserView(null)
  attachedWin = null
  attachedId = null
  views.clear()
  viewUrls.clear()
}
