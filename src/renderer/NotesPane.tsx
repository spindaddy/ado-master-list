import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { LocalNote, NowPin } from '../../shared/types'
import { NOW_PIN_LIMIT, nowPinKey } from '../../shared/types'

function newNote(): LocalNote {
  const now = new Date().toISOString()
  return {
    id: `note-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`,
    title: '',
    body: '',
    done: false,
    createdAt: now,
    updatedAt: now
  }
}

export function NotesPane(props: {
  onOpenCount?: (count: number) => void
  focusId?: string | null
  pins?: NowPin[]
  onTogglePin?: (id: string) => void
}) {
  const [notes, setNotes] = useState<LocalNote[]>([])
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [filter, setFilter] = useState('')
  const [showDone, setShowDone] = useState(false)
  const saveTimer = useRef<number | null>(null)
  const notesRef = useRef<LocalNote[]>([])

  const persist = useCallback((next: LocalNote[]) => {
    notesRef.current = next
    setNotes(next)
    props.onOpenCount?.(next.filter((n) => !n.done).length)
    if (saveTimer.current) window.clearTimeout(saveTimer.current)
    saveTimer.current = window.setTimeout(() => {
      void window.adoApi.saveLocalNotes(next)
    }, 280)
  }, [props.onOpenCount])

  useEffect(() => {
    void window.adoApi.getLocalNotes().then((list) => {
      notesRef.current = list
      setNotes(list)
      props.onOpenCount?.(list.filter((n) => !n.done).length)
      setSelectedId((current) => current ?? list.find((n) => !n.done)?.id ?? list[0]?.id ?? null)
    })
    return () => {
      if (saveTimer.current) window.clearTimeout(saveTimer.current)
      void window.adoApi.saveLocalNotes(notesRef.current)
    }
  }, [props.onOpenCount])

  useEffect(() => {
    if (props.focusId) setSelectedId(props.focusId)
  }, [props.focusId])

  const visible = useMemo(() => {
    const q = filter.trim().toLowerCase()
    return notes.filter((n) => {
      if (!showDone && n.done) return false
      if (!q) return true
      return `${n.title} ${n.body}`.toLowerCase().includes(q)
    })
  }, [notes, filter, showDone])

  const selected = notes.find((n) => n.id === selectedId) ?? null
  const openCount = notes.filter((n) => !n.done).length

  const patchSelected = (patch: Partial<LocalNote>) => {
    if (!selected) return
    persist(
      notes.map((n) =>
        n.id === selected.id
          ? { ...n, ...patch, updatedAt: new Date().toISOString() }
          : n
      )
    )
  }

  const addNote = () => {
    const note = newNote()
    persist([note, ...notes])
    setSelectedId(note.id)
  }

  const removeSelected = () => {
    if (!selected) return
    const next = notes.filter((n) => n.id !== selected.id)
    persist(next)
    setSelectedId(next.find((n) => !n.done)?.id ?? next[0]?.id ?? null)
  }

  return (
    <div className="layout">
      <section className="list-pane">
        <div className="pane-label-row">
          <div className="pane-label">Local notes</div>
          <input
            className="search"
            placeholder="Filter…"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
          />
        </div>
        <div className="notes-list-actions">
          <button className="btn btn-primary" type="button" onClick={addNote}>
            Add note
          </button>
          <label className="setup-active notes-show-done">
            <input
              type="checkbox"
              checked={showDone}
              onChange={(e) => setShowDone(e.target.checked)}
            />
            <span>Show done ({notes.length - openCount})</span>
          </label>
        </div>
        <div className="item-list">
          {visible.length === 0 ? (
            <p className="sidebar-empty">
              {notes.length === 0
                ? 'Add a note for tasks that are not in Azure DevOps.'
                : 'No notes match this filter.'}
            </p>
          ) : (
            visible.map((note) => (
              <button
                key={note.id}
                type="button"
                className={`item-row${selectedId === note.id ? ' selected' : ''}${note.done ? ' note-done' : ''}`}
                onClick={() => setSelectedId(note.id)}
              >
                <div className="item-row-top">
                  <span className="item-row-title">
                    {note.title.trim() || 'Untitled note'}
                  </span>
                  {(props.pins ?? []).some(
                    (p) => nowPinKey(p) === nowPinKey({ kind: 'note', id: note.id })
                  ) ? (
                    <span className="chip">Now</span>
                  ) : null}
                </div>
                <div className="item-row-meta">
                  {note.done ? 'Done' : 'Open'}
                  {note.body.trim()
                    ? ` · ${note.body.trim().slice(0, 80)}`
                    : ''}
                </div>
              </button>
            ))
          )}
        </div>
      </section>
      <section className="detail-pane">
        {!selected ? (
          <p className="detail-empty">Select a note or add one.</p>
        ) : (
          <>
            <div className="detail-header">
              <div>
                <div className="detail-kicker">Local · not in ADO</div>
                <h1>{selected.title.trim() || 'Untitled note'}</h1>
              </div>
              <div className="row-actions">
                {props.onTogglePin ? (
                  <button
                    className="btn"
                    type="button"
                    onClick={() => props.onTogglePin?.(selected.id)}
                  >
                    {(props.pins ?? []).some(
                      (p) => nowPinKey(p) === nowPinKey({ kind: 'note', id: selected.id })
                    )
                      ? 'Unpin from Now'
                      : (props.pins ?? []).length >= NOW_PIN_LIMIT
                        ? 'Now is full'
                        : 'Pin to Now'}
                  </button>
                ) : null}
                <label className="setup-active">
                  <input
                    type="checkbox"
                    checked={selected.done}
                    onChange={(e) => patchSelected({ done: e.target.checked })}
                  />
                  <span>Done</span>
                </label>
                <button className="btn btn-danger" type="button" onClick={removeSelected}>
                  Delete
                </button>
              </div>
            </div>
            <div className="field">
              <label htmlFor="note-title">Title</label>
              <input
                id="note-title"
                value={selected.title}
                placeholder="What needs doing?"
                onChange={(e) => patchSelected({ title: e.target.value })}
              />
            </div>
            <div className="field">
              <label htmlFor="note-body">Notes</label>
              <textarea
                className="tall"
                id="note-body"
                rows={16}
                value={selected.body}
                placeholder="Details, links, reminders…"
                onChange={(e) => patchSelected({ body: e.target.value })}
              />
            </div>
          </>
        )}
      </section>
    </div>
  )
}
