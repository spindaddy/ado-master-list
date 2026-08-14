import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { MS_WEB_EMBEDS } from './msWebEmbedMeta'

export function MsWebEmbed(props: {
  id: string
  title: string
  url?: string
  hidden: boolean
}) {
  const hostRef = useRef<HTMLDivElement>(null)
  const [alertsMuted, setAlertsMuted] = useState(false)
  const url =
    props.url ||
    (props.id === 'outlook' || props.id === 'teams'
      ? MS_WEB_EMBEDS[props.id].url
      : '')

  useEffect(() => {
    if (props.id !== 'teams') return
    void window.adoApi.getAlertsMuted().then(setAlertsMuted)
    return window.adoApi.onAlertsMuted(setAlertsMuted)
  }, [props.id])

  const syncBounds = useCallback(() => {
    const el = hostRef.current
    if (props.hidden || !el) {
      void window.adoApi.hideMsWeb()
      return
    }
    const rect = el.getBoundingClientRect()
    if (rect.width < 8 || rect.height < 8) {
      void window.adoApi.hideMsWeb()
      return
    }
    void window.adoApi.showMsWeb({
      id: props.id,
      url: url || undefined,
      bounds: {
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height
      }
    })
  }, [props.hidden, props.id, url])

  useLayoutEffect(() => {
    syncBounds()
    const el = hostRef.current
    const observer = el ? new ResizeObserver(() => syncBounds()) : null
    if (el && observer) observer.observe(el)
    window.addEventListener('resize', syncBounds)
    return () => {
      observer?.disconnect()
      window.removeEventListener('resize', syncBounds)
      void window.adoApi.hideMsWeb()
    }
  }, [syncBounds])

  return (
    <section className="day-cal">
      <div className="day-cal-toolbar">
        <h2>{props.title}</h2>
        <div className="day-cal-nav">
          {props.id === 'teams' ? (
            <button
              className={alertsMuted ? 'btn btn-mute-on' : 'btn'}
              type="button"
              title={
                alertsMuted
                  ? 'Alerts are off. Click to turn ticket, Outlook, and Teams sounds back on.'
                  : 'Silence ticket, Outlook, and Teams alerts while you are on a Teams call.'
              }
              onClick={() => void window.adoApi.setAlertsMuted(!alertsMuted)}
            >
              {alertsMuted ? 'Muted · on call' : 'Mute for call'}
            </button>
          ) : null}
          <button
            className="btn"
            type="button"
            onClick={() => {
              void window.adoApi.reloadMsWeb(props.id, url || undefined)
              syncBounds()
            }}
          >
            Reload
          </button>
          {url ? (
            <button
              className="btn"
              type="button"
              onClick={() => void window.adoApi.openExternal(url)}
            >
              Open in browser
            </button>
          ) : null}
        </div>
      </div>
      <div
        className="outlook-web-host"
        ref={hostRef}
        onMouseDown={() => void window.adoApi.focusMsWeb()}
      />
    </section>
  )
}
