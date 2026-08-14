import {
  BrowserView,
  BrowserWindow,
  desktopCapturer,
  shell,
  systemPreferences,
  type Session,
  type WebContents
} from 'electron'
import { attachEditContextMenu } from './editMenu'

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
  const PREFIX = '${NOTIFY_PREFIX}';
  const report = (title, options) => {
    try {
      console.log(PREFIX + JSON.stringify({
        title: String(title || ''),
        body: String((options && options.body) || ''),
        tag: String((options && options.tag) || '')
      }));
    } catch (e) {}
  };
  const Orig = window.Notification;
  if (Orig && !Orig.__adoMasterHook) {
    const Wrapped = function Notification(title, options) {
      report(title, options);
      try { return new Orig(title, options); } catch (e) { return undefined; }
    };
    Wrapped.__adoMasterHook = true;
    Wrapped.prototype = Orig.prototype;
    Object.defineProperty(Wrapped, 'permission', {
      configurable: true,
      get: function () { return Orig.permission; }
    });
    Wrapped.requestPermission = function () {
      return Orig.requestPermission.apply(Orig, arguments);
    };
    try { Wrapped.maxActions = Orig.maxActions; } catch (e) {}
    window.Notification = Wrapped;
    try { Orig.requestPermission(); } catch (e) {}
  }
  try {
    const proto = ServiceWorkerRegistration && ServiceWorkerRegistration.prototype;
    if (proto && proto.showNotification && !proto.__adoMasterHook) {
      const origShow = proto.showNotification;
      proto.showNotification = function (title, options) {
        report(title, options);
        return origShow.apply(this, arguments);
      };
      proto.__adoMasterHook = true;
    }
  } catch (e) {}
})();`

const PRESENCE_JS = `(() => {
  const visible = () => 'visible';
  try {
    Object.defineProperty(Document.prototype, 'hidden', { configurable: true, get: () => false });
    Object.defineProperty(Document.prototype, 'visibilityState', { configurable: true, get: visible });
  } catch (e) {}
  try {
    Object.defineProperty(document, 'hidden', { configurable: true, get: () => false });
    Object.defineProperty(document, 'visibilityState', { configurable: true, get: visible });
  } catch (e) {}
  if (window.__adoMasterPresence) return;
  window.__adoMasterPresence = true;
  document.addEventListener('visibilitychange', (e) => { e.stopImmediatePropagation(); }, true);
  window.addEventListener('blur', (e) => { e.stopImmediatePropagation(); }, true);
  try { Document.prototype.hasFocus = function () { return true; }; } catch (e) {}
  try { document.hasFocus = function () { return true; }; } catch (e) {}
  const poke = () => {
    try {
      const x = 6 + Math.floor(Math.random() * 20);
      const y = 6 + Math.floor(Math.random() * 20);
      document.dispatchEvent(new MouseEvent('mousemove', {
        bubbles: true, cancelable: true, view: window, clientX: x, clientY: y
      }));
    } catch (e) {}
  };
  poke();
  setInterval(poke, 40000);
})();`

const views = new Map<MsWebEmbedId, BrowserView>()
let attachedWin: BrowserWindow | null = null
let attachedId: MsWebEmbedId | null = null
const hookedSessions = new WeakSet<Session>()
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

const OUTLOOK_UNREAD_JS = `(() => {
  const unreadFrom = (text) => {
    const t = String(text || '');
    if (!/unread/i.test(t)) return 0;
    if (/calendar|meeting|event|reminder/i.test(t) && !/inbox|mail|folder/i.test(t)) return 0;
    const m =
      t.match(/(\\d+)\\s+unread/i) ||
      t.match(/unread[^0-9]{0,20}(\\d+)/i);
    if (!m) return 0;
    const n = Number(m[1]);
    return Number.isFinite(n) && n >= 0 && n < 10000 ? n : 0;
  };
  let best = unreadFrom(document.title || '');
  const nodes = document.querySelectorAll('[aria-label], [title]');
  for (const el of nodes) {
    const text = [el.getAttribute('aria-label') || '', el.getAttribute('title') || ''].join(' ');
    best = Math.max(best, unreadFrom(text));
  }
  return best;
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
  const code = id === 'outlook' ? OUTLOOK_UNREAD_JS : COUNT_JS
  const readFrame = async (exec: (js: string) => Promise<unknown>): Promise<number> => {
    try {
      const value = await exec(code)
      const n = Number(value)
      return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0
    } catch {
      return 0
    }
  }
  let best = 0
  try {
    const frames = view.webContents.mainFrame?.framesInSubtree
    if (frames && frames.length > 0) {
      for (const frame of frames) {
        best = Math.max(best, await readFrame((js) => frame.executeJavaScript(js, true)))
      }
    } else {
      best = await readFrame((js) => view.webContents.executeJavaScript(js, true))
    }
  } catch {
    best = await readFrame((js) => view.webContents.executeJavaScript(js, true))
  }
  return best
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
  ensureView('teams')
  if (countTimer) return
  countTimer = setInterval(() => {
    const teams = views.get('teams')
    if (teams && !teams.webContents.isDestroyed()) {
      injectPageHooks('teams', teams.webContents)
    }
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

function hookSession(sess: Session): void {
  if (hookedSessions.has(sess)) return
  hookedSessions.add(sess)
  sess.setPermissionRequestHandler((_wc, permission, callback) => {
    callback(
      permission === 'media' ||
        permission === 'display-capture' ||
        permission === 'notifications' ||
        permission === 'clipboard-sanitized-write' ||
        permission === 'clipboard-read'
    )
  })
  sess.setPermissionCheckHandler((_wc, permission) => {
    return (
      permission === 'media' ||
      permission === 'display-capture' ||
      permission === 'notifications' ||
      permission === 'clipboard-sanitized-write' ||
      permission === 'clipboard-read'
    )
  })
  sess.setDisplayMediaRequestHandler(
    async (request, callback) => {
      try {
        const sources = await desktopCapturer.getSources({
          types: ['screen', 'window'],
          thumbnailSize: { width: 0, height: 0 }
        })
        const screen = sources.find((s) => s.id.startsWith('screen:')) ?? sources[0]
        if (!screen) {
          callback({})
          return
        }
        callback({
          video: screen,
          ...(request.audioRequested ? { audio: 'loopback' as const } : {})
        })
      } catch (err) {
        console.error('Screen share failed', err)
        callback({})
      }
    },
    { useSystemPicker: true }
  )
}

async function requestCallMediaAccess(): Promise<void> {
  if (process.platform !== 'darwin') return
  try {
    await systemPreferences.askForMediaAccess('microphone')
    await systemPreferences.askForMediaAccess('camera')
  } catch {
    // ignore
  }
}

function injectPageHooks(id: MsWebEmbedId, wc: WebContents): void {
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
  if (id === 'teams') run(PRESENCE_JS)
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
  if (id === 'outlook') return
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

function preventHtmlFullscreen(wc: WebContents): void {
  wc.on('enter-html-full-screen', () => {
    void wc.executeJavaScript(
      'document.fullscreenElement && document.exitFullscreen && document.exitFullscreen()',
      true
    ).catch(() => undefined)
    if (attachedWin && !attachedWin.isDestroyed() && attachedWin.isFullScreen()) {
      attachedWin.setFullScreen(false)
    }
  })
}

function detachBrowserViews(win: BrowserWindow): void {
  try {
    const list = typeof win.getBrowserViews === 'function' ? win.getBrowserViews() : []
    if (list.length > 0) {
      for (const view of list) win.removeBrowserView(view)
      return
    }
  } catch {
    // fall through
  }
  win.setBrowserView(null)
}

function attachBrowserView(win: BrowserWindow, view: BrowserView): void {
  try {
    if (typeof win.addBrowserView === 'function') {
      const list = win.getBrowserViews()
      for (const existing of list) {
        if (existing !== view) win.removeBrowserView(existing)
      }
      if (!win.getBrowserViews().includes(view)) win.addBrowserView(view)
      if (typeof win.setTopBrowserView === 'function') win.setTopBrowserView(view)
      return
    }
  } catch {
    // fall through
  }
  win.setBrowserView(view)
}

function consoleMessageText(...args: unknown[]): string {
  if (typeof args[2] === 'string') return args[2]
  const first = args[0] as { message?: string } | undefined
  if (first && typeof first.message === 'string') return first.message
  const second = args[1] as { message?: string } | undefined
  if (second && typeof second.message === 'string') return second.message
  return ''
}

function wireView(id: MsWebEmbedId, view: BrowserView): void {
  hookSession(view.webContents.session)
  attachEditContextMenu(view.webContents)
  preventHtmlFullscreen(view.webContents)
  view.webContents.setBackgroundThrottling(false)
  view.webContents.setUserAgent(CHROME_UA)
  view.webContents.on('dom-ready', () => injectPageHooks(id, view.webContents))
  view.webContents.on('did-finish-load', () => {
    injectPageHooks(id, view.webContents)
    void refreshCounts()
  })
  view.webContents.on('did-frame-finish-load', () => injectPageHooks(id, view.webContents))
  view.webContents.on('did-navigate-in-page', () => injectPageHooks(id, view.webContents))
  view.webContents.on('page-title-updated', (_e, title) => {
    applyTitleCount(id, title)
  })
  view.webContents.on('console-message', (...args: unknown[]) => {
    const parsed = parseNotify(consoleMessageText(...args))
    if (!parsed) return
    notifyHandler?.({ source: id, title: parsed.title, body: parsed.body })
    void refreshCounts()
  })
  view.webContents.on('did-create-window', (child) => {
    hookSession(child.webContents.session)
    attachEditContextMenu(child.webContents)
    preventHtmlFullscreen(child.webContents)
    injectPageHooks(id, child.webContents)
    child.webContents.setBackgroundThrottling(false)
    child.webContents.on('did-finish-load', () => injectPageHooks(id, child.webContents))
    child.webContents.on('console-message', (...args: unknown[]) => {
      const parsed = parseNotify(consoleMessageText(...args))
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
  if (id === 'teams') void requestCallMediaAccess()
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

function toWindowBounds(
  win: BrowserWindow,
  bounds: { x: number; y: number; width: number; height: number }
): { x: number; y: number; width: number; height: number } {
  const windowBounds = win.getBounds()
  const contentBounds = win.getContentBounds()
  const dx = contentBounds.x - windowBounds.x
  const dy = contentBounds.y - windowBounds.y
  return {
    x: Math.max(0, Math.round(bounds.x + dx)),
    y: Math.max(0, Math.round(bounds.y + dy)),
    width: Math.max(1, Math.round(bounds.width)),
    height: Math.max(1, Math.round(bounds.height))
  }
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
    detachBrowserViews(attachedWin)
  }
  attachBrowserView(win, next)
  attachedWin = win
  attachedId = id
  next.setBounds(toWindowBounds(win, bounds))
}

export function hideMsWebEmbed(win: BrowserWindow | null): void {
  if (!win || win.isDestroyed()) return
  detachBrowserViews(win)
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
      detachBrowserViews(attachedWin)
      attachedId = null
    }
    views.delete(id)
    viewUrls.delete(id)
  }
}

export function disposeMsWebEmbeds(): void {
  if (attachedWin && !attachedWin.isDestroyed()) detachBrowserViews(attachedWin)
  attachedWin = null
  attachedId = null
  views.clear()
  viewUrls.clear()
}
