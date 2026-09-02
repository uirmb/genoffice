import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'

export const MARKDOWN_ZOOM_MIN = 30
export const MARKDOWN_ZOOM_MAX = 400
export const MARKDOWN_ZOOM_STEP = 10
export const MARKDOWN_ZOOM_DEFAULT = 100

function clampZoom(value: number): number {
  return Math.min(MARKDOWN_ZOOM_MAX, Math.max(MARKDOWN_ZOOM_MIN, value))
}

function labels(): {
  out: string
  slider: string
  reset: string
  in: string
} {
  const chinese = document.documentElement.lang.toLowerCase().startsWith('zh')
  return chinese
    ? { out: '缩小', slider: '文档缩放', reset: '恢复 100%', in: '放大' }
    : { out: 'Zoom out', slider: 'Document zoom', reset: 'Reset to 100%', in: 'Zoom in' }
}

/**
 * The status bar belongs to App, while zoom is purely view state and must never
 * affect Markdown serialization. Portal the control into that chrome and apply
 * CSS zoom only to the paper surface.
 */
export function ZoomControl() {
  const [statusBar, setStatusBar] = useState<HTMLElement | null>(null)
  const [zoom, setZoom] = useState(MARKDOWN_ZOOM_DEFAULT)

  useEffect(() => {
    const current = document.querySelector<HTMLElement>('.status-bar')
    if (current) {
      setStatusBar(current)
      return
    }

    const observer = new MutationObserver(() => {
      const next = document.querySelector<HTMLElement>('.status-bar')
      if (!next) return
      setStatusBar(next)
      observer.disconnect()
    })
    observer.observe(document.getElementById('root') ?? document.body, {
      childList: true,
      subtree: true,
    })
    return () => observer.disconnect()
  }, [])

  useEffect(() => {
    if (!statusBar) return
    const page = document.querySelector<HTMLElement>('.doc-page')
    if (!page) return
    page.style.setProperty('zoom', String(zoom / 100))
    page.dataset.zoom = String(zoom)
  }, [statusBar, zoom])

  if (!statusBar) return null

  const text = labels()
  const set = (value: number) => setZoom(clampZoom(value))

  return createPortal(
    <div className="status-right markdown-zoom-control" aria-label={text.slider}>
      <button
        type="button"
        className="zoom-btn markdown-zoom-out"
        title={text.out}
        aria-label={text.out}
        disabled={zoom <= MARKDOWN_ZOOM_MIN}
        onClick={() => set(zoom - MARKDOWN_ZOOM_STEP)}
      >
        −
      </button>
      <input
        className="zoom-slider markdown-zoom-slider"
        type="range"
        min={MARKDOWN_ZOOM_MIN}
        max={MARKDOWN_ZOOM_MAX}
        step={MARKDOWN_ZOOM_STEP}
        value={zoom}
        aria-label={text.slider}
        onChange={(event) => set(Number(event.currentTarget.value))}
      />
      <button
        type="button"
        className="zoom-value markdown-zoom-value"
        title={text.reset}
        aria-label={text.reset}
        onClick={() => set(MARKDOWN_ZOOM_DEFAULT)}
      >
        {zoom}%
      </button>
      <button
        type="button"
        className="zoom-btn markdown-zoom-in"
        title={text.in}
        aria-label={text.in}
        disabled={zoom >= MARKDOWN_ZOOM_MAX}
        onClick={() => set(zoom + MARKDOWN_ZOOM_STEP)}
      >
        +
      </button>
    </div>,
    statusBar,
  )
}
