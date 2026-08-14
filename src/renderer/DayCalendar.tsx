import { useMemo } from 'react'
import type { OutlookEventDto } from '../../shared/types'

const HOUR_PX = 64

export function toLocalYmd(date: Date): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  return `${y}-${m}-${d}`
}

export function addDays(date: Date, days: number): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days)
}

export function parseOutlookDate(raw: string): Date {
  if (!raw) return new Date(NaN)
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [y, m, d] = raw.split('-').map(Number)
    return new Date(y, m - 1, d)
  }
  if (/[zZ]$/.test(raw) || /[+-]\d{2}:\d{2}$/.test(raw)) return new Date(raw)
  return new Date(`${raw}Z`)
}

function minutesOnDay(date: Date): number {
  return date.getHours() * 60 + date.getMinutes()
}

function formatHour(hour: number): string {
  const d = new Date(2000, 0, 1, hour, 0, 0)
  return d.toLocaleTimeString(undefined, { hour: 'numeric' })
}

type LaidOut = {
  event: OutlookEventDto
  col: number
  cols: number
  top: number
  height: number
}

function layoutTimed(
  events: OutlookEventDto[],
  dayStart: Date,
  dayEnd: Date,
  startHour: number
): LaidOut[] {
  const items = events
    .map((event) => {
      const start = parseOutlookDate(event.start)
      const end = parseOutlookDate(event.end)
      const clampedStart = Math.max(start.getTime(), dayStart.getTime())
      const clampedEnd = Math.min(end.getTime(), dayEnd.getTime())
      const startMin = minutesOnDay(new Date(clampedStart))
      const endMin = Math.max(
        startMin + 20,
        minutesOnDay(new Date(clampedEnd - 1)) + 1
      )
      return {
        event,
        startMin,
        endMin,
        top: ((startMin - startHour * 60) / 60) * HOUR_PX,
        height: Math.max(28, ((endMin - startMin) / 60) * HOUR_PX)
      }
    })
    .sort((a, b) => a.startMin - b.startMin || a.endMin - b.endMin)

  const colEnds: number[] = []
  const withCol = items.map((item) => {
    let col = colEnds.findIndex((end) => end <= item.startMin)
    if (col < 0) {
      col = colEnds.length
      colEnds.push(item.endMin)
    } else {
      colEnds[col] = item.endMin
    }
    return { ...item, col }
  })

  return withCol.map((item) => {
    let cols = item.col + 1
    for (const other of withCol) {
      const overlap = other.startMin < item.endMin && other.endMin > item.startMin
      if (overlap) cols = Math.max(cols, other.col + 1)
    }
    return {
      event: item.event,
      col: item.col,
      cols,
      top: item.top,
      height: item.height
    }
  })
}

export function DayCalendar(props: {
  ymd: string
  events: OutlookEventDto[]
  loading: boolean
  error: string | null
  connected: boolean
  onPrev: () => void
  onNext: () => void
  onToday: () => void
  onOpenSettings: () => void
}) {
  const todayYmd = toLocalYmd(new Date())
  const isToday = props.ymd === todayYmd
  const [year, month, day] = props.ymd.split('-').map(Number)
  const dayStart = new Date(year, month - 1, day, 0, 0, 0, 0)
  const dayEnd = new Date(year, month - 1, day + 1, 0, 0, 0, 0)
  const heading = dayStart.toLocaleDateString(undefined, {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric'
  })

  const allDay = props.events.filter((e) => e.isAllDay)
  const timed = props.events.filter((e) => !e.isAllDay)

  const { startHour, hours } = useMemo(() => {
    let start = 6
    let end = 21
    for (const event of timed) {
      const s = parseOutlookDate(event.start)
      const e = parseOutlookDate(event.end)
      if (!Number.isNaN(s.getTime())) start = Math.min(start, s.getHours())
      if (!Number.isNaN(e.getTime())) {
        const hour = e.getMinutes() > 0 || e.getSeconds() > 0 ? e.getHours() + 1 : e.getHours()
        end = Math.max(end, hour)
      }
    }
    start = Math.max(0, start)
    end = Math.min(24, Math.max(start + 1, end))
    return {
      startHour: start,
      hours: Array.from({ length: end - start }, (_, i) => start + i)
    }
  }, [timed])

  const laidOut = useMemo(
    () => layoutTimed(timed, dayStart, dayEnd, startHour),
    [timed, dayStart.getTime(), dayEnd.getTime(), startHour]
  )

  const nowTop = useMemo(() => {
    if (!isToday) return null
    const now = new Date()
    return ((minutesOnDay(now) - startHour * 60) / 60) * HOUR_PX
  }, [isToday, startHour, props.events.length])

  return (
    <section className="day-cal">
      <div className="day-cal-toolbar">
        <div className="day-cal-nav">
          <button className="btn" type="button" onClick={props.onPrev} aria-label="Previous day">
            ‹
          </button>
          <button className="btn" type="button" onClick={props.onNext} aria-label="Next day">
            ›
          </button>
          {!isToday && (
            <button className="btn" type="button" onClick={props.onToday}>
              Today
            </button>
          )}
        </div>
        <h2>{heading}</h2>
        <div className="day-cal-count">
          {props.loading ? 'Loading…' : `${props.events.length} event${props.events.length === 1 ? '' : 's'}`}
        </div>
      </div>

      {!props.connected ? (
        <div className="detail-empty">
          Sign in to Outlook in Settings to see this day’s calendar.
          <div style={{ marginTop: '0.75rem' }}>
            <button className="btn btn-primary" type="button" onClick={props.onOpenSettings}>
              Open Settings
            </button>
          </div>
        </div>
      ) : props.error ? (
        <div className="detail-empty">{props.error}</div>
      ) : (
        <div className="day-cal-scroll">
          {allDay.length > 0 && (
            <div className="day-cal-allday">
              <div className="day-cal-allday-label">All day</div>
              <div className="day-cal-allday-items">
                {allDay.map((event) => (
                  <button
                    key={`${event.id}-${event.start}`}
                    type="button"
                    className="day-cal-chip"
                    onClick={() => void window.adoApi.openExternal(event.webLink)}
                  >
                    {event.subject}
                  </button>
                ))}
              </div>
            </div>
          )}

          <div
            className="day-cal-grid"
            style={{ height: hours.length * HOUR_PX }}
          >
            {hours.map((hour, i) => (
              <div
                key={hour}
                className="day-cal-hour"
                style={{ top: i * HOUR_PX, height: HOUR_PX }}
              >
                <span>{formatHour(hour)}</span>
              </div>
            ))}

            <div className="day-cal-events">
              {laidOut.map((item) => (
                <button
                  key={`${item.event.id}-${item.event.start}`}
                  type="button"
                  className="day-cal-event"
                  style={{
                    top: item.top,
                    height: item.height,
                    left: `calc(${(item.col / item.cols) * 100}% + 4px)`,
                    width: `calc(${100 / item.cols}% - 8px)`
                  }}
                  onClick={() => void window.adoApi.openExternal(item.event.webLink)}
                >
                  <strong>{item.event.subject}</strong>
                  {item.event.location ? <span>{item.event.location}</span> : null}
                </button>
              ))}
            </div>

            {nowTop != null && nowTop >= 0 && nowTop <= hours.length * HOUR_PX && (
              <div className="day-cal-now" style={{ top: nowTop }}>
                <i />
              </div>
            )}
          </div>

          {!props.loading && props.events.length === 0 && (
            <p className="sidebar-empty">Nothing on the calendar this day.</p>
          )}
        </div>
      )}
    </section>
  )
}
