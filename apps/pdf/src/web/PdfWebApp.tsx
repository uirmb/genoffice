import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, ReactElement, RefObject } from 'react'
import { GlobalWorkerOptions, TextLayer, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import type {
  OfficeDocumentKind,
  OfficeFile,
  OfficeFileDescriptor,
  OfficeHostApi,
} from '@genoffice/office-host-api'
import { OFFICE_PROTOCOL_VERSION } from '@genoffice/office-protocol'
import type { EditorIframeBridge, WebRuntimeMode } from '@genoffice/web-runtime'
import { OutlinePanel } from '../renderer/OutlinePanel'
import type { OutlineNode } from '../renderer/OutlinePanel'
import { buildSearchIndex, searchInIndex } from '../renderer/search'
import type { SearchIndex, SearchMatch } from '../renderer/search'
import { isPdfOfficeFile, PDF_WEB_ACCEPT } from './viewer-policy'

GlobalWorkerOptions.workerSrc = workerUrl

const ASSET_BASE = new URL('pdfjs/', document.baseURI).href
const DOC_OPTS = {
  cMapUrl: `${ASSET_BASE}cmaps/`,
  cMapPacked: true,
  standardFontDataUrl: `${ASSET_BASE}standard_fonts/`,
  wasmUrl: `${ASSET_BASE}wasm/`,
}
const ZOOM_STEPS = [0.5, 0.67, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4]

interface PageMetric {
  width: number
  height: number
}

interface PreparedPdf {
  file: OfficeFile
  doc: PDFDocumentProxy
  pageMetrics: PageMetric[]
  outline: OutlineNode[]
}

interface RawLinkAnnotation {
  subtype?: string
  rect?: number[]
  url?: string
  unsafeUrl?: string
  dest?: unknown
}

interface PdfWebAppProps {
  host: OfficeHostApi
  bridge: EditorIframeBridge | null
  runtimeMode: WebRuntimeMode
}

function descriptorWithBytes(descriptor: OfficeFileDescriptor, source: OfficeFile): OfficeFile {
  return {
    ...source,
    ...descriptor,
    bytes: source.bytes,
    transport: 'buffer',
  }
}

function clampScale(scale: number): number {
  return Math.min(4, Math.max(0.25, scale))
}

function useVisibleSet(
  rootRef: RefObject<HTMLElement | null>,
  count: number,
  rootMargin: string,
): { visible: Set<number>; setItemRef: (index: number) => (element: HTMLElement | null) => void } {
  const [visible, setVisible] = useState<Set<number>>(new Set())
  const itemRefs = useRef<(HTMLElement | null)[]>([])

  useEffect(() => {
    const root = rootRef.current
    if (!root || count === 0) return
    const observer = new IntersectionObserver(
      (entries) => {
        setVisible((previous) => {
          const next = new Set(previous)
          for (const entry of entries) {
            const index = Number((entry.target as HTMLElement).dataset.index)
            if (entry.isIntersecting) next.add(index)
            else next.delete(index)
          }
          return next
        })
      },
      { root, rootMargin },
    )
    for (const element of itemRefs.current) if (element) observer.observe(element)
    return () => observer.disconnect()
  }, [count, rootMargin, rootRef])

  return {
    visible,
    setItemRef: (index) => (element) => {
      itemRefs.current[index] = element
    },
  }
}

function PdfPage({
  doc,
  pageNo,
  scale,
  visible,
  searchTarget,
  onGoToDest,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  scale: number
  visible: boolean
  searchTarget: boolean
  onGoToDest: (dest: unknown) => void
}): ReactElement {
  const holderRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const holder = holderRef.current
    if (!holder || !visible) return
    let cancelled = false
    let renderTask: RenderTask | null = null

    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const viewport = page.getViewport({ scale })
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.floor(viewport.width * dpr))
      canvas.height = Math.max(1, Math.floor(viewport.height * dpr))
      canvas.style.width = `${Math.floor(viewport.width)}px`
      canvas.style.height = `${Math.floor(viewport.height)}px`

      renderTask = page.render({
        canvas,
        viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      })
      try {
        await renderTask.promise
      } catch {
        return
      }
      if (cancelled) return

      const textLayerElement = document.createElement('div')
      textLayerElement.className = 'textLayer'
      const linkLayerElement = document.createElement('div')
      linkLayerElement.className = 'pdf-web-link-layer'
      holder.replaceChildren(canvas, textLayerElement, linkLayerElement)

      const textLayer = new TextLayer({
        textContentSource: page.streamTextContent(),
        container: textLayerElement,
        viewport,
      })
      try {
        await textLayer.render()
      } catch {
        // The page may have left the viewport while the text layer was rendering.
      }
      if (cancelled) return

      const annotations = (await page.getAnnotations({ intent: 'display' })) as RawLinkAnnotation[]
      if (cancelled) return
      for (const annotation of annotations) {
        if (annotation.subtype !== 'Link' || !annotation.rect) continue
        const url = annotation.url || annotation.unsafeUrl
        if (!url && annotation.dest == null) continue
        const rect = viewport.convertToViewportRectangle(annotation.rect)
        const left = Math.min(rect[0], rect[2])
        const top = Math.min(rect[1], rect[3])
        const anchor = document.createElement('a')
        anchor.className = 'pdf-web-link'
        anchor.style.left = `${left}px`
        anchor.style.top = `${top}px`
        anchor.style.width = `${Math.abs(rect[2] - rect[0])}px`
        anchor.style.height = `${Math.abs(rect[3] - rect[1])}px`
        if (url) {
          anchor.href = url
          anchor.target = '_blank'
          anchor.rel = 'noreferrer'
          anchor.title = url
        } else {
          anchor.href = '#'
          anchor.addEventListener('click', (event) => {
            event.preventDefault()
            onGoToDest(annotation.dest)
          })
        }
        linkLayerElement.append(anchor)
      }
    })()

    return () => {
      cancelled = true
      renderTask?.cancel()
      holder.replaceChildren()
    }
  }, [doc, onGoToDest, pageNo, scale, visible])

  return (
    <div
      ref={holderRef}
      className={`pdf-web-page-content${searchTarget ? ' pdf-web-page-search-target' : ''}`}
    />
  )
}

function PdfThumbnail({
  doc,
  pageNo,
  visible,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  visible: boolean
}): ReactElement {
  const canvasRef = useRef<HTMLCanvasElement>(null)

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas || !visible) return
    let cancelled = false
    let renderTask: RenderTask | null = null

    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const base = page.getViewport({ scale: 1 })
      const scale = Math.min(1, 112 / base.width)
      const viewport = page.getViewport({ scale })
      const dpr = Math.min(window.devicePixelRatio || 1, 1.5)
      canvas.width = Math.max(1, Math.floor(viewport.width * dpr))
      canvas.height = Math.max(1, Math.floor(viewport.height * dpr))
      canvas.style.width = `${Math.floor(viewport.width)}px`
      canvas.style.height = `${Math.floor(viewport.height)}px`
      renderTask = page.render({
        canvas,
        viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      })
      try {
        await renderTask.promise
      } catch {
        // cancelled
      }
    })()

    return () => {
      cancelled = true
      renderTask?.cancel()
    }
  }, [doc, pageNo, visible])

  return <canvas ref={canvasRef} />
}

export function PdfWebApp({ host, bridge, runtimeMode }: PdfWebAppProps): ReactElement {
  const [currentFile, setCurrentFile] = useState<OfficeFile | null>(null)
  const [pdfDocument, setPdfDocument] = useState<PDFDocumentProxy | null>(null)
  const [pageMetrics, setPageMetrics] = useState<PageMetric[]>([])
  const [outline, setOutline] = useState<OutlineNode[]>([])
  const [sidebar, setSidebar] = useState<'thumbs' | 'outline' | null>('thumbs')
  const [currentPage, setCurrentPage] = useState(1)
  const [scale, setScale] = useState(1)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [searchMatches, setSearchMatches] = useState<SearchMatch[]>([])
  const [searchMatchIndex, setSearchMatchIndex] = useState(-1)
  const [searching, setSearching] = useState(false)
  const scrollRef = useRef<HTMLDivElement>(null)
  const thumbsRef = useRef<HTMLDivElement>(null)
  const searchIndexRef = useRef<SearchIndex | null>(null)
  const prepareSequenceRef = useRef(0)

  const pageCount = pdfDocument?.numPages ?? 0
  const currentSearchPage =
    searchMatchIndex >= 0 ? (searchMatches[searchMatchIndex]?.pageIndex ?? -1) + 1 : -1
  const pages = useMemo(() => Array.from({ length: pageCount }, (_, index) => index + 1), [pageCount])
  const pageVisibility = useVisibleSet(scrollRef, pageCount, '900px 0px')
  const thumbVisibility = useVisibleSet(thumbsRef, pageCount, '500px 0px')

  const preparePdf = useCallback(async (file: OfficeFile): Promise<PreparedPdf> => {
    if (!isPdfOfficeFile(file)) throw new Error(`Unsupported file type: ${file.name}`)
    const task = getDocument({ data: file.bytes.slice(0), ...DOC_OPTS })
    const doc = await task.promise
    try {
      const metrics: PageMetric[] = []
      for (let pageNo = 1; pageNo <= doc.numPages; pageNo += 1) {
        const page = await doc.getPage(pageNo)
        const viewport = page.getViewport({ scale: 1 })
        metrics.push({ width: viewport.width, height: viewport.height })
      }
      const rawOutline = await doc.getOutline()
      return {
        file,
        doc,
        pageMetrics: metrics,
        outline: (rawOutline ?? []) as OutlineNode[],
      }
    } catch (prepareError) {
      await doc.destroy()
      throw prepareError
    }
  }, [])

  const installPreparedPdf = useCallback(
    (prepared: PreparedPdf, authoritativeFile?: OfficeFileDescriptor): void => {
      const nextFile = authoritativeFile
        ? descriptorWithBytes(authoritativeFile, prepared.file)
        : prepared.file
      setPdfDocument((previous) => {
        if (previous && previous !== prepared.doc) void previous.destroy()
        return prepared.doc
      })
      setCurrentFile(nextFile)
      setPageMetrics(prepared.pageMetrics)
      setOutline(prepared.outline)
      setCurrentPage(1)
      setScale(1)
      setQuery('')
      setSearchMatches([])
      setSearchMatchIndex(-1)
      searchIndexRef.current = null
      setError(null)
      host.setDirty(false)
      host.setTitle(nextFile.name)
    },
    [host],
  )

  const loadInitialFile = useCallback(
    async (file: OfficeFile): Promise<void> => {
      const sequence = ++prepareSequenceRef.current
      setLoading(true)
      setError(null)
      try {
        const prepared = await preparePdf(file)
        if (sequence !== prepareSequenceRef.current) {
          await prepared.doc.destroy()
          return
        }
        installPreparedPdf(prepared)
      } catch (loadError) {
        if (sequence === prepareSequenceRef.current) {
          setError(loadError instanceof Error ? loadError.message : String(loadError))
        }
      } finally {
        if (sequence === prepareSequenceRef.current) setLoading(false)
      }
    },
    [installPreparedPdf, preparePdf],
  )

  const openPdf = useCallback(async (): Promise<void> => {
    if (loading) return
    setLoading(true)
    setError(null)
    try {
      if (host.pickDocument) {
        const selection = await host.pickDocument({ accept: [...PDF_WEB_ACCEPT] })
        if (selection.status === 'cancelled') return
        if (selection.status === 'failed') throw new Error(selection.error)

        let prepared: PreparedPdf | null = null
        try {
          prepared = await preparePdf(selection.file)
          const bound = host.confirmDocumentOpened
            ? await host.confirmDocumentOpened(selection.selectionId)
            : { ok: true as const }
          if (!bound.ok) {
            throw new Error(bound.error || 'The selected PDF could not be bound as current.')
          }
          installPreparedPdf(prepared, bound.file)
          prepared = null
        } catch (openError) {
          if (prepared) await prepared.doc.destroy()
          await host.releasePickedDocument?.(selection.selectionId)
          throw openError
        }
        return
      }

      const selected = await host.pickFile({
        accept: [...PDF_WEB_ACCEPT],
        multiple: false,
        mode: 'file',
      })
      const first = selected?.[0]
      if (!first) return
      const file =
        first.transport === 'buffer' && first.bytes
          ? ({ ...first, bytes: first.bytes, transport: 'buffer' } as OfficeFile)
          : await host.readFile(first.id)
      const prepared = await preparePdf(file)
      installPreparedPdf(prepared)
    } catch (openError) {
      setError(openError instanceof Error ? openError.message : String(openError))
    } finally {
      setLoading(false)
    }
  }, [host, installPreparedPdf, loading, preparePdf])

  const scrollToPage = useCallback((pageNo: number): void => {
    const normalized = Math.min(Math.max(pageNo, 1), pageCount || 1)
    document.getElementById(`pdf-web-page-${normalized}`)?.scrollIntoView({ block: 'start' })
    setCurrentPage(normalized)
  }, [pageCount])

  const goToDestination = useCallback(
    (dest: unknown): void => {
      if (!pdfDocument || dest == null) return
      void (async () => {
        try {
          const resolved = typeof dest === 'string' ? await pdfDocument.getDestination(dest) : dest
          if (!Array.isArray(resolved) || resolved.length === 0) return
          const ref = resolved[0]
          if (!ref || typeof ref !== 'object') return
          const pageIndex = await pdfDocument.getPageIndex(ref as Parameters<PDFDocumentProxy['getPageIndex']>[0])
          scrollToPage(pageIndex + 1)
        } catch {
          // Invalid destinations are ignored, matching normal PDF viewer behavior.
        }
      })()
    },
    [pdfDocument, scrollToPage],
  )

  const fitWidth = useCallback((): void => {
    const container = scrollRef.current
    const metric = pageMetrics[currentPage - 1]
    if (!container || !metric) return
    setScale(clampScale((container.clientWidth - 48) / metric.width))
  }, [currentPage, pageMetrics])

  const fitPage = useCallback((): void => {
    const container = scrollRef.current
    const metric = pageMetrics[currentPage - 1]
    if (!container || !metric) return
    const widthScale = (container.clientWidth - 48) / metric.width
    const heightScale = (container.clientHeight - 32) / metric.height
    setScale(clampScale(Math.min(widthScale, heightScale)))
  }, [currentPage, pageMetrics])

  const changeZoom = useCallback((direction: -1 | 1): void => {
    setScale((current) => {
      const index = ZOOM_STEPS.findIndex((value) => value >= current - 0.001)
      const nextIndex = Math.min(
        ZOOM_STEPS.length - 1,
        Math.max(0, (index < 0 ? 4 : index) + direction),
      )
      return ZOOM_STEPS[nextIndex] ?? current
    })
  }, [])

  const runSearch = useCallback(async (): Promise<void> => {
    const normalized = query.trim()
    if (!pdfDocument || !normalized) {
      setSearchMatches([])
      setSearchMatchIndex(-1)
      return
    }
    setSearching(true)
    try {
      if (!searchIndexRef.current) searchIndexRef.current = await buildSearchIndex(pdfDocument)
      const matches = searchInIndex(searchIndexRef.current, normalized)
      setSearchMatches(matches)
      setSearchMatchIndex(matches.length ? 0 : -1)
      if (matches[0]) scrollToPage(matches[0].pageIndex + 1)
    } finally {
      setSearching(false)
    }
  }, [pdfDocument, query, scrollToPage])

  const stepSearch = useCallback(
    (direction: -1 | 1): void => {
      if (!searchMatches.length) return
      const next = (searchMatchIndex + direction + searchMatches.length) % searchMatches.length
      setSearchMatchIndex(next)
      const match = searchMatches[next]
      if (match) scrollToPage(match.pageIndex + 1)
    },
    [scrollToPage, searchMatchIndex, searchMatches],
  )

  const handleScroll = useCallback((): void => {
    const container = scrollRef.current
    if (!container) return
    const center = container.getBoundingClientRect().top + container.clientHeight / 2
    let bestPage = currentPage
    let bestDistance = Number.POSITIVE_INFINITY
    for (const element of container.querySelectorAll<HTMLElement>('[data-pdf-page]')) {
      const rect = element.getBoundingClientRect()
      const distance = Math.abs(rect.top + rect.height / 2 - center)
      if (distance < bestDistance) {
        bestDistance = distance
        bestPage = Number(element.dataset.pdfPage)
      }
    }
    if (Number.isFinite(bestPage) && bestPage !== currentPage) setCurrentPage(bestPage)
  }, [currentPage])

  useEffect(() => {
    if (!bridge) return
    const unsubscribe = bridge.subscribe((message) => {
      switch (message.type) {
        case 'office:init':
          if ((message.payload.kind as string) !== 'pdf') {
            bridge.send({
              protocol: OFFICE_PROTOCOL_VERSION,
              type: 'office:error',
              requestId: message.requestId,
              payload: { code: 'UNSUPPORTED_KIND', message: 'PDF Web only accepts PDF files.' },
            })
            return
          }
          void loadInitialFile(message.payload.file)
          return
        case 'office:new':
          bridge.send({
            protocol: OFFICE_PROTOCOL_VERSION,
            type: 'office:error',
            requestId: message.requestId,
            payload: { code: 'READ_ONLY', message: 'PDF Web is viewer-only and cannot create PDFs.' },
          })
          return
        case 'office:set-mode':
          return
        case 'office:set-locale':
          document.documentElement.lang = message.payload.locale
          return
        case 'office:save':
          bridge.send({
            protocol: OFFICE_PROTOCOL_VERSION,
            type: 'office:save-result',
            requestId: message.requestId,
            payload: { ok: false, error: 'PDF Web is viewer-only.' },
          })
          return
        case 'office:query-state':
          bridge.send({
            protocol: OFFICE_PROTOCOL_VERSION,
            type: 'office:state-result',
            requestId: message.requestId,
            payload: {
              ready: true,
              dirty: false,
              saving: false,
              mode: 'view',
              title: currentFile?.name,
            },
          })
          return
        case 'office:request-close':
          void host.approveClose?.(message.requestId)
          return
        default:
          return
      }
    })

    bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:ready',
      payload: { kind: 'pdf' as OfficeDocumentKind },
    })
    return unsubscribe
  }, [bridge, currentFile?.name, host, loadInitialFile])

  useEffect(() => {
    host.setDirty(false)
    return () => {
      prepareSequenceRef.current += 1
      if (pdfDocument) void pdfDocument.destroy()
    }
  }, [host, pdfDocument])

  const pageStyle = useCallback(
    (pageNo: number): CSSProperties => {
      const metric = pageMetrics[pageNo - 1]
      if (!metric) return {}
      return {
        width: Math.max(1, Math.round(metric.width * scale)),
        height: Math.max(1, Math.round(metric.height * scale)),
        '--scale-factor': String(scale),
      } as CSSProperties
    },
    [pageMetrics, scale],
  )

  return (
    <div className="pdf-web-app">
      <header className="pdf-web-toolbar">
        <div className="pdf-web-product">PDF</div>
        <button className="pdf-web-button primary" type="button" onClick={() => void openPdf()} disabled={loading}>
          打开
        </button>
        <div className="pdf-web-separator" />
        <button className="pdf-web-button compact" type="button" onClick={() => scrollToPage(currentPage - 1)} disabled={!pdfDocument || currentPage <= 1} aria-label="上一页">
          ‹
        </button>
        <label className="pdf-web-page-field">
          <input
            value={pdfDocument ? currentPage : ''}
            inputMode="numeric"
            aria-label="当前页"
            onChange={(event) => {
              const value = Number(event.target.value)
              if (Number.isFinite(value) && value > 0) setCurrentPage(Math.min(value, pageCount || 1))
            }}
            onKeyDown={(event) => {
              if (event.key === 'Enter') scrollToPage(currentPage)
            }}
            disabled={!pdfDocument}
          />
          <span>/ {pageCount || 0}</span>
        </label>
        <button className="pdf-web-button compact" type="button" onClick={() => scrollToPage(currentPage + 1)} disabled={!pdfDocument || currentPage >= pageCount} aria-label="下一页">
          ›
        </button>
        <div className="pdf-web-separator" />
        <button className="pdf-web-button compact" type="button" onClick={() => changeZoom(-1)} disabled={!pdfDocument} aria-label="缩小">
          −
        </button>
        <span className="pdf-web-zoom">{Math.round(scale * 100)}%</span>
        <button className="pdf-web-button compact" type="button" onClick={() => changeZoom(1)} disabled={!pdfDocument} aria-label="放大">
          +
        </button>
        <button className="pdf-web-button" type="button" onClick={fitWidth} disabled={!pdfDocument}>
          适合宽度
        </button>
        <button className="pdf-web-button" type="button" onClick={fitPage} disabled={!pdfDocument}>
          适合页面
        </button>
        <div className="pdf-web-toolbar-spacer" />
        <form
          className="pdf-web-search"
          onSubmit={(event) => {
            event.preventDefault()
            void runSearch()
          }}
        >
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="搜索文档"
            aria-label="搜索文档"
            disabled={!pdfDocument}
          />
          <button type="submit" disabled={!pdfDocument || searching}>查找</button>
          <span className="pdf-web-search-count">
            {searchMatches.length && searchMatchIndex >= 0
              ? `${searchMatchIndex + 1}/${searchMatches.length}`
              : searchMatches.length === 0 && query && !searching
                ? '0'
                : ''}
          </span>
          <button type="button" onClick={() => stepSearch(-1)} disabled={!searchMatches.length} aria-label="上一个匹配">
            ↑
          </button>
          <button type="button" onClick={() => stepSearch(1)} disabled={!searchMatches.length} aria-label="下一个匹配">
            ↓
          </button>
        </form>
      </header>

      <div className="pdf-web-main">
        {pdfDocument && sidebar ? (
          <aside className="pdf-web-sidebar">
            <div className="pdf-web-sidebar-tabs">
              <button className={sidebar === 'thumbs' ? 'active' : ''} type="button" onClick={() => setSidebar('thumbs')}>
                页面
              </button>
              <button className={sidebar === 'outline' ? 'active' : ''} type="button" onClick={() => setSidebar('outline')}>
                目录
              </button>
              <button type="button" className="close" onClick={() => setSidebar(null)} aria-label="关闭侧栏">
                ×
              </button>
            </div>
            {sidebar === 'thumbs' ? (
              <div ref={thumbsRef} className="pdf-web-thumbs">
                {pages.map((pageNo, index) => (
                  <button
                    key={pageNo}
                    ref={thumbVisibility.setItemRef(index) as (element: HTMLButtonElement | null) => void}
                    data-index={index}
                    type="button"
                    className={`pdf-web-thumb${currentPage === pageNo ? ' active' : ''}`}
                    onClick={() => scrollToPage(pageNo)}
                  >
                    <span className="pdf-web-thumb-canvas">
                      <PdfThumbnail doc={pdfDocument} pageNo={pageNo} visible={thumbVisibility.visible.has(index)} />
                    </span>
                    <span>{pageNo}</span>
                  </button>
                ))}
              </div>
            ) : (
              <div className="pdf-web-outline">
                {outline.length ? (
                  <OutlinePanel outline={outline} onGoToDest={goToDestination} />
                ) : (
                  <div className="pdf-web-empty-side">此 PDF 没有目录</div>
                )}
              </div>
            )}
          </aside>
        ) : null}

        {pdfDocument && !sidebar ? (
          <button className="pdf-web-sidebar-open" type="button" onClick={() => setSidebar('thumbs')}>
            页面
          </button>
        ) : null}

        <main ref={scrollRef} className="pdf-web-scroll" onScroll={handleScroll}>
          {!pdfDocument ? (
            <div className="pdf-web-empty">
              <div className="pdf-web-empty-card">
                <div className="pdf-web-empty-mark">PDF</div>
                <h1>PDF 预览</h1>
                <p>Web 版本仅提供阅读能力，不包含编辑、批注、签名或表单修改。</p>
                <button type="button" onClick={() => void openPdf()} disabled={loading}>
                  {loading ? '正在打开…' : runtimeMode === 'embedded' ? '从系统中选择 PDF' : '打开 PDF'}
                </button>
              </div>
            </div>
          ) : (
            <div className="pdf-web-pages">
              {pages.map((pageNo, index) => (
                <section
                  id={`pdf-web-page-${pageNo}`}
                  key={pageNo}
                  ref={pageVisibility.setItemRef(index) as (element: HTMLElement | null) => void}
                  data-index={index}
                  data-pdf-page={pageNo}
                  className={`pdf-web-page${currentSearchPage === pageNo ? ' search-target' : ''}`}
                  style={pageStyle(pageNo)}
                >
                  <PdfPage
                    doc={pdfDocument}
                    pageNo={pageNo}
                    scale={scale}
                    visible={pageVisibility.visible.has(index)}
                    searchTarget={currentSearchPage === pageNo}
                    onGoToDest={goToDestination}
                  />
                </section>
              ))}
            </div>
          )}
        </main>
      </div>

      <footer className="pdf-web-statusbar">
        <span>{currentFile?.name ?? '未打开 PDF'}</span>
        <span className="pdf-web-status-spacer" />
        <span>只读</span>
        {pdfDocument ? <span>{pageCount} 页</span> : null}
      </footer>

      {loading && pdfDocument ? <div className="pdf-web-loading">正在打开 PDF…</div> : null}
      {error ? (
        <div className="pdf-web-error" role="alert">
          <span>{error}</span>
          <button type="button" onClick={() => setError(null)}>关闭</button>
        </div>
      ) : null}
    </div>
  )
}
