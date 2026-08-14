import { useEffect, useState } from 'react'
import type { LocalNote, MasterEntry, NowPin } from '../../shared/types'
import { NOW_PIN_LIMIT, nowPinKey } from '../../shared/types'

export function NowPane(props: {
  master: MasterEntry[]
  pins: NowPin[]
  refreshToken?: number
  onUnpin: (pin: NowPin) => void
  onOpenWork: (id: string) => void
  onOpenNote: (id: string) => void
  onNoteDone: (id: string, done: boolean) => void
}) {
  const [notes, setNotes] = useState<LocalNote[]>([])

  useEffect(() => {
    void window.adoApi.getLocalNotes().then(setNotes)
  }, [props.pins, props.refreshToken])

  const slots = Array.from({ length: NOW_PIN_LIMIT }, (_, i) => props.pins[i] ?? null)

  return (
    <div className="now-page">
      <div className="now-intro">
        <h2>Now</h2>
        <p>
          At most three things. Pin from Work or Notes. Everything else can wait.
        </p>
      </div>
      <div className="now-grid">
        {slots.map((pin, index) => {
          if (!pin) {
            return (
              <div key={`empty-${index}`} className="now-card now-card-empty">
                <div className="now-slot">Slot {index + 1}</div>
                <p>Pin a work item or a local note. {NOW_PIN_LIMIT - props.pins.length} left.</p>
              </div>
            )
          }
          if (pin.kind === 'work') {
            const entry = props.master.find((e) => e.id === pin.id)
            return (
              <article key={nowPinKey(pin)} className="now-card">
                <div className="now-slot">ADO · slot {index + 1}</div>
                {entry ? (
                  <>
                    <h3>{entry.title}</h3>
                    <p className="now-meta">
                      #{entry.workItemId} · {entry.organization} / {entry.project} ·{' '}
                      {entry.state}
                    </p>
                    <div className="row-actions">
                      <button
                        className="btn btn-primary"
                        type="button"
                        onClick={() => props.onOpenWork(entry.id)}
                      >
                        Open in Work
                      </button>
                      <button
                        className="btn"
                        type="button"
                        onClick={() => void window.adoApi.openExternal(entry.webUrl)}
                      >
                        Open in ADO
                      </button>
                      <button className="btn btn-ghost" type="button" onClick={() => props.onUnpin(pin)}>
                        Unpin
                      </button>
                    </div>
                  </>
                ) : (
                  <>
                    <h3>Item left the list</h3>
                    <p className="now-meta">It may have been unassigned or closed.</p>
                    <button className="btn" type="button" onClick={() => props.onUnpin(pin)}>
                      Unpin
                    </button>
                  </>
                )}
              </article>
            )
          }
          const note = notes.find((n) => n.id === pin.id)
          return (
            <article key={nowPinKey(pin)} className="now-card">
              <div className="now-slot">Note · slot {index + 1}</div>
              {note ? (
                <>
                  <h3>{note.title.trim() || 'Untitled note'}</h3>
                  <p className="now-meta">
                    {note.done ? 'Done' : 'Open'}
                    {note.body.trim() ? ` · ${note.body.trim().slice(0, 160)}` : ''}
                  </p>
                  <div className="row-actions">
                    <button
                      className="btn btn-primary"
                      type="button"
                      onClick={() => props.onOpenNote(note.id)}
                    >
                      Open in Notes
                    </button>
                    <button
                      className="btn"
                      type="button"
                      onClick={() => props.onNoteDone(note.id, !note.done)}
                    >
                      {note.done ? 'Not done' : 'Mark done'}
                    </button>
                    <button className="btn btn-ghost" type="button" onClick={() => props.onUnpin(pin)}>
                      Unpin
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <h3>Note missing</h3>
                  <p className="now-meta">It was deleted from Notes.</p>
                  <button className="btn" type="button" onClick={() => props.onUnpin(pin)}>
                    Unpin
                  </button>
                </>
              )}
            </article>
          )
        })}
      </div>
    </div>
  )
}
