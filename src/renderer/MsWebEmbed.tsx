import { useCallback, useLayoutEffect, useRef } from 'react'
import { MS_WEB_EMBEDS } from './msWebEmbedMeta'

export function MsWebEmbed(props: {
  id: string
  title: string
  url?: string
  hidden: boolean
}) {
  const hostRef = useRef<HTMLDivElement>(null)
  const url =
    props.url ||
    (props.id === 'outlook' || props.id === 'teams'
      ? MS_WEB_EMBEDS[props.id].url
      : '')

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
