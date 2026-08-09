import { normalizeLang, type Lang } from '@genoffice/i18n'
import type {
  OfficeFile,
  OfficeFileDescriptor,
  OfficeHostApi,
  SaveDocumentInput,
  SelectedOfficeFile,
} from '@genoffice/office-host-api'
import {
  OFFICE_PROTOCOL_VERSION,
  type HostToEditorMessage,
} from '@genoffice/office-protocol'
import {
  addElement,
  addPicture,
  addSection,
  builtinLayoutInfos,
  BUILTIN_LAYOUT_PREFIX,
  commitSaved,
  createBlankPptx,
  deleteElement,
  deleteSlide,
  duplicateSlide,
  EMU_PER_PT,
  ensureBuiltinLayout,
  ensureRunLinkRels,
  getSections,
  getSlideComments,
  getSlideNotes,
  getSlideTransition,
  insertBlankSlide,
  insertSlideWithLayout,
  listSlideLayouts,
  materializeSlide,
  moveSection,
  moveSlide as moveSlideModel,
  openPptx,
  removeSection,
  renameSection,
  resetSlideLayout,
  savePptx,
  setElementFont,
  setElementParagraphFormat,
  setSections,
  setSlideBackground,
  setSlideHidden,
  setSlideLayout,
  setSlideNotes,
  setSlideSize,
  shouldOfferBuiltinLayouts,
  updateConnectorsForMoved,
  type OpenedPptx,
  type Paragraph,
  type Slide,
  type TextElement,
} from '@genoffice/pptx-engine'
import { buildRenderSlide, EMU_PER_PX_96, type RenderSlide } from '@genoffice/pptx-render'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import { applyEditParagraphs, collectParagraphFormatPatches, levelsChanged } from '../main/edit-text'
import type { MenuCommand, OpenResult, SlidesApi } from '../shared/ipc'

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const IMAGE_MIMES = [
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/bmp',
  'image/webp',
  'image/svg+xml',
] as const

interface WebSession {
  opened: OpenedPptx
  fitWidthPx: number
  undoStack: Uint8Array[]
  redoStack: Uint8Array[]
  transformPreviewSnapshot: Uint8Array | null
}

function bytesToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

function arrayBufferToBase64(bytes: ArrayBuffer): string {
  const src = new Uint8Array(bytes)
  const chunk = 0x8000
  let binary = ''
  for (let i = 0; i < src.length; i += chunk) {
    binary += String.fromCharCode(...src.subarray(i, Math.min(i + chunk, src.length)))
  }
  return btoa(binary)
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value)
  const out = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i)
  return out
}

function mimeForPath(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase()
  if (ext === 'png') return 'image/png'
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg'
  if (ext === 'gif') return 'image/gif'
  if (ext === 'bmp') return 'image/bmp'
  if (ext === 'webp') return 'image/webp'
  if (ext === 'svg') return 'image/svg+xml'
  return 'application/octet-stream'
}

function mediaResolver(opened: OpenedPptx): (mediaRef: string) => string | undefined {
  return (mediaRef) => {
    if (/^https?:\/\//i.test(mediaRef) || /^data:/i.test(mediaRef)) return mediaRef
    const bytes = opened.archive.readBytes(mediaRef)
    if (!bytes) return undefined
    return `data:${mimeForPath(mediaRef)};base64,${arrayBufferToBase64(bytesToArrayBuffer(bytes))}`
  }
}

function buildSlide(opened: OpenedPptx, index: number, fitWidthPx: number): RenderSlide | null {
  const slide = opened.deck.slides[index]
  if (!slide) return null
  return buildRenderSlide(slide, opened.deck.size, {
    fitWidthPx,
    media: mediaResolver(opened),
    slideNo: index + 1,
  })
}

function buildAllSlides(opened: OpenedPptx, fitWidthPx: number): RenderSlide[] {
  return opened.deck.slides
    .map((_, index) => buildSlide(opened, index, fitWidthPx))
    .filter((slide): slide is RenderSlide => slide !== null)
}

function virtualPath(file: OfficeFileDescriptor): string {
  return `web-office://files/${encodeURIComponent(file.id)}/${encodeURIComponent(file.name)}`
}

async function selectedToOfficeFile(host: OfficeHostApi, selected: SelectedOfficeFile): Promise<OfficeFile> {
  if (selected.transport === 'buffer' && selected.bytes) {
    return {
      id: selected.id,
      name: selected.name,
      mimeType: selected.mimeType,
      size: selected.size ?? selected.bytes.byteLength,
      version: selected.version ?? null,
      bytes: selected.bytes.slice(0),
    }
  }
  return host.readFile(selected.id)
}

function extensionOf(file: OfficeFile): string {
  const ext = file.name.split('.').pop()?.toLowerCase()
  if (ext) return ext === 'jpeg' ? 'jpg' : ext
  if (file.mimeType === 'image/png') return 'png'
  if (file.mimeType === 'image/jpeg') return 'jpg'
  if (file.mimeType === 'image/gif') return 'gif'
  if (file.mimeType === 'image/bmp') return 'bmp'
  if (file.mimeType === 'image/webp') return 'webp'
  if (file.mimeType === 'image/svg+xml') return 'svg'
  return ''
}

function isImageMime(mime: string): boolean {
  return IMAGE_MIMES.includes(mime as (typeof IMAGE_MIMES)[number])
}

function safeName(name: string | undefined, fallback: string): string {
  const value = (name || fallback).trim()
  return /\.pptx$/i.test(value) ? value : `${value}.pptx`
}

export interface SlidesWebController {
  slidesApi: SlidesApi
  destroy(): void
}

export function createSlidesWebController(
  host: OfficeHostApi,
  bridge?: EditorIframeBridge,
): SlidesWebController {
  let session: WebSession | null = null
  let currentFile: OfficeFileDescriptor | null = null
  let currentLang: Lang = normalizeLang(document.documentElement.lang || navigator.language || 'en')
  let mode: 'view' | 'edit' = 'edit'
  let dirty = false
  let saving = false
  let readyNotified = false
  let initialOpenResolve: ((result: OpenResult | null) => void) | null = null
  let lastFitWidthPx = 960
  const pendingSaveRequestIds = new Set<string>()

  const languageHandlers = new Set<(lang: Lang) => void>()
  const modeHandlers = new Set<(mode: 'view' | 'edit') => void>()
  const menuHandlers = new Set<(command: MenuCommand) => void>()
  const openedHandlers = new Set<(result: OpenResult) => void>()
  const renamedHandlers = new Set<(path: string) => void>()

  const setDirtyState = (next: boolean) => {
    if (dirty === next) return
    dirty = next
    host.setDirty(next)
  }

  const setMode = (next: 'view' | 'edit') => {
    if (mode === next) return
    mode = next
    document.documentElement.classList.toggle('office-view-mode', mode === 'view')
    for (const handler of modeHandlers) handler(mode)
  }

  const setLanguage = (locale: string) => {
    const next = normalizeLang(locale)
    if (next === currentLang) return
    currentLang = next
    for (const handler of languageHandlers) handler(next)
  }

  const title = () => currentFile?.name ?? 'Untitled.pptx'

  const openResult = (): OpenResult | null => {
    if (!session) return null
    return {
      path: currentFile ? virtualPath(currentFile) : '',
      slides: buildAllSlides(session.opened, session.fitWidthPx),
      size: { ...session.opened.deck.size },
    }
  }

  const replaceOpenedFromBytes = async (
    bytes: ArrayBuffer,
    fitWidthPx: number,
    file: OfficeFileDescriptor | null,
  ): Promise<OpenResult> => {
    const opened = await openPptx(new Uint8Array(bytes))
    session = {
      opened,
      fitWidthPx,
      undoStack: [],
      redoStack: [],
      transformPreviewSnapshot: null,
    }
    currentFile = file ? { ...file, size: bytes.byteLength } : null
    lastFitWidthPx = fitWidthPx
    setDirtyState(false)
    host.setTitle(currentFile?.name ?? 'Untitled Presentation')
    return openResult()!
  }

  const openOfficeFile = async (file: OfficeFile, fitWidthPx: number): Promise<OpenResult> =>
    replaceOpenedFromBytes(
      file.bytes.slice(0),
      fitWidthPx,
      {
        id: file.id,
        name: file.name,
        mimeType: file.mimeType || PPTX_MIME,
        size: file.size ?? file.bytes.byteLength,
        version: file.version ?? null,
      },
    )

  const openSelected = async (fitWidthPx: number): Promise<OpenResult | null> => {
    const selected = await host.pickFile({
      multiple: false,
      accept: [PPTX_MIME, '.pptx'],
      mode: 'file',
    })
    if (!selected?.[0]) return null
    return openOfficeFile(await selectedToOfficeFile(host, selected[0]), fitWidthPx)
  }

  const snapshot = async (): Promise<Uint8Array | null> => {
    if (!session) return null
    return (await savePptx(session.opened)).slice()
  }

  const pushHistory = async (): Promise<void> => {
    if (!session) return
    const bytes = await snapshot()
    if (!bytes) return
    session.undoStack.push(bytes)
    if (session.undoStack.length > 50) session.undoStack.shift()
    session.redoStack = []
  }

  const restoreSnapshot = async (bytes: Uint8Array): Promise<RenderSlide[]> => {
    if (!session) return []
    session.opened = await openPptx(bytes)
    session.transformPreviewSnapshot = null
    setDirtyState(true)
    return buildAllSlides(session.opened, session.fitWidthPx)
  }

  const reportPendingHostSave = (ok: boolean, error?: string) => {
    if (!bridge || pendingSaveRequestIds.size === 0) return
    for (const requestId of pendingSaveRequestIds) {
      bridge.send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-result',
        requestId,
        payload: { ok, error },
      })
    }
    pendingSaveRequestIds.clear()
  }

  const persist = async (
    saveMode: 'save' | 'saveAs',
    defaultName?: string,
  ): Promise<{ ok: boolean; path?: string; error?: string; slides?: RenderSlide[] }> => {
    if (!session) return { ok: false, error: 'No presentation is open.' }
    if (mode === 'view') return { ok: false, error: 'The presentation is read-only.' }

    const bytes = await savePptx(session.opened)
    const newDocument = currentFile === null
    const fallback: OfficeFileDescriptor = currentFile ?? {
      id: `new:${Date.now()}`,
      name: safeName(defaultName, 'Untitled.pptx'),
      mimeType: PPTX_MIME,
      size: bytes.byteLength,
      version: null,
    }
    const file: OfficeFileDescriptor = {
      ...fallback,
      name: safeName(defaultName ?? fallback.name, 'Untitled.pptx'),
      mimeType: PPTX_MIME,
      size: bytes.byteLength,
    }

    const input: SaveDocumentInput = {
      file,
      bytes: bytesToArrayBuffer(bytes),
      baseVersion: currentFile?.version ?? null,
      mode: saveMode,
      newDocument: newDocument || undefined,
    }

    saving = true
    try {
      const result = await host.saveDocument(input)
      if (!result.ok) {
        reportPendingHostSave(false, result.error)
        return { ok: false, error: result.error || 'Save failed.' }
      }
      currentFile = {
        ...(result.file ?? file),
        mimeType: result.file?.mimeType || PPTX_MIME,
        size: bytes.byteLength,
      }
      commitSaved(session.opened)
      session.undoStack = []
      session.redoStack = []
      setDirtyState(false)
      host.setTitle(currentFile.name)
      const resultOpen = openResult()!
      reportPendingHostSave(true)
      return { ok: true, path: resultOpen.path, slides: resultOpen.slides }
    } finally {
      saving = false
    }
  }

  const saveHistoryVersion = async (): Promise<{ ok: boolean; error?: string }> => {
    if (!session || !currentFile) {
      return { ok: false, error: 'Save the new presentation before creating history.' }
    }
    if (!host.saveHistoryVersion) {
      return { ok: false, error: 'History versions are not supported by this host.' }
    }
    const bytes = await savePptx(session.opened)
    const result = await host.saveHistoryVersion({
      file: currentFile,
      bytes: bytesToArrayBuffer(bytes),
      baseVersion: currentFile.version ?? null,
    })
    return { ok: result.ok, error: result.error }
  }

  const exportPptx = async (): Promise<{ ok: boolean; error?: string }> => {
    if (!session) return { ok: false, error: 'No presentation is open.' }
    if (!host.exportDocument) return { ok: false, error: 'PPTX export is not supported by this host.' }
    const bytes = await savePptx(session.opened)
    const descriptor: OfficeFileDescriptor = currentFile ?? {
      id: `export:${Date.now()}`,
      name: 'Untitled.pptx',
      mimeType: PPTX_MIME,
      size: bytes.byteLength,
      version: null,
    }
    const result = await host.exportDocument({
      format: 'pptx',
      file: { ...descriptor, name: safeName(descriptor.name, 'Untitled.pptx'), size: bytes.byteLength },
      bytes: bytesToArrayBuffer(bytes),
    })
    return { ok: result.ok, error: result.error }
  }

  const requireSlide = (index: number): Slide | null => session?.opened.deck.slides[index] ?? null

  const rebuild = (index: number): RenderSlide | null =>
    session ? buildSlide(session.opened, index, session.fitWidthPx) : null

  const toEmuFor = (fitWidthPx: number) => {
    if (!session) return (_px: number) => 0
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = fitWidthPx / baseWidthPx
    return (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
  }

  const explicit: Record<PropertyKey, unknown> = {
    getLanguage: async () => currentLang,
    onLanguageChanged: (handler: (lang: Lang) => void) => {
      languageHandlers.add(handler)
      return () => languageHandlers.delete(handler)
    },
    getHostEditorMode: async () => mode,
    onHostEditorModeChanged: (handler: (next: 'view' | 'edit') => void) => {
      modeHandlers.add(handler)
      return () => modeHandlers.delete(handler)
    },
    reportDirtyChange: (next: boolean) => setDirtyState(next),
    openPptx: (fitWidthPx: number) => openSelected(fitWidthPx),
    openPptxPath: async (path: string, fitWidthPx: number) => {
      if (!currentFile || virtualPath(currentFile) !== path || !session) return null
      session.fitWidthPx = fitWidthPx
      lastFitWidthPx = fitWidthPx
      return openResult()
    },
    consumePendingOpen: async (fitWidthPx: number) => {
      lastFitWidthPx = fitWidthPx
      if (!bridge) return null
      if (!readyNotified) {
        readyNotified = true
        bridge.send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:ready',
          payload: { kind: 'pptx' },
        })
      }
      return new Promise<OpenResult | null>((resolve) => {
        initialOpenResolve = resolve
      })
    },
    newBlank: async (fitWidthPx: number) => {
      const bytes = await createBlankPptx()
      return replaceOpenedFromBytes(bytesToArrayBuffer(bytes), fitWidthPx, null)
    },
    getRenderSlides: async () => (session ? buildAllSlides(session.opened, session.fitWidthPx) : null),
    getSlideSize: async () => (session ? { ...session.opened.deck.size } : null),
    getLayouts: async () => {
      if (!session) return null
      const layouts = listSlideLayouts(session.opened.archive)
      if (shouldOfferBuiltinLayouts(layouts)) {
        layouts.push(
          ...builtinLayoutInfos(session.opened.deck.size, new Set(layouts.map((layout) => layout.name))),
        )
      }
      return { layouts, size: { ...session.opened.deck.size } }
    },
    editText: async (op: any) => {
      if (mode !== 'edit' || !session || op.groupId) return null
      const slide = requireSlide(op.slideIndex)
      const el = slide?.elements.find((item) => item.id === op.sourceId) as TextElement | undefined
      if (!slide || !el?.text) return null
      await pushHistory()
      const levelDirty = levelsChanged(el.text.paragraphs, op.paragraphs)
      el.text.paragraphs = applyEditParagraphs(el.text.paragraphs, op.paragraphs)
      ensureRunLinkRels(session.opened, op.slideIndex, el.text.paragraphs)
      el.dirty = true
      for (const { index, patch } of collectParagraphFormatPatches(op.paragraphs)) {
        setElementParagraphFormat(slide, op.sourceId, patch, [index])
      }
      if (levelDirty) {
        el.dirtyPPr = { ...el.dirtyPPr, level: true, indents: true }
        materializeSlide(session.opened, op.slideIndex)
      }
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    setElementFont: async (op: any) => {
      if (mode !== 'edit') return null
      const slide = requireSlide(op.slideIndex)
      if (!slide) return null
      await pushHistory()
      let changed = false
      for (const id of op.sourceIds ?? []) {
        if (
          setElementFont(slide, id, {
            fontFamily: op.fontFamily,
            fontSizePt: op.fontSizePt,
            strike: op.strike,
            bold: op.bold,
            italic: op.italic,
            underline: op.underline,
            color: op.color,
          })
        ) {
          changed = true
        }
      }
      if (!changed) return null
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    setElementParagraphFormat: async (op: any) => {
      if (mode !== 'edit') return null
      const slide = requireSlide(op.slideIndex)
      if (!slide) return null
      await pushHistory()
      let changed = false
      const patch = {
        bullet: op.bullet,
        bulletChar: op.bulletChar,
        bulletHangEmu: op.bulletHangEmu,
        bulletSizePct: op.bulletSizePct,
        bulletColor: op.bulletColor,
        lineSpacingPct: op.lineSpacingPct,
        spaceBeforePt: op.spaceBeforePt,
        spaceAfterPt: op.spaceAfterPt,
        align: op.align,
        indentDelta: op.indentDelta,
      }
      for (const id of op.sourceIds ?? []) {
        if (setElementParagraphFormat(slide, id, patch)) changed = true
      }
      if (!changed) return null
      if (op.indentDelta) materializeSlide(session!.opened, op.slideIndex)
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    editTransform: async (op: any) => {
      if (mode !== 'edit' || !session || op.groupId) return null
      const slide = requireSlide(op.slideIndex)
      const el = slide?.elements.find((item) => item.id === op.sourceId)
      if (!slide || !el) return null
      if (op.preview) {
        if (!session.transformPreviewSnapshot) session.transformPreviewSnapshot = await snapshot()
      } else if (!session.transformPreviewSnapshot) {
        await pushHistory()
      } else {
        session.undoStack.push(session.transformPreviewSnapshot)
        session.redoStack = []
        session.transformPreviewSnapshot = null
      }
      const toEmu = toEmuFor(op.fitWidthPx)
      el.transform = {
        ...el.transform,
        offset: {
          x: toEmu(op.xPx),
          y: toEmu(op.yPx),
          cx: toEmu(op.wPx),
          cy: toEmu(op.hPx),
        },
        rot: Math.round(op.rotationDeg * 60000),
      }
      el.dirtyTransform = true
      updateConnectorsForMoved(slide, [op.sourceId])
      session.fitWidthPx = op.fitWidthPx
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    batchEditTransform: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      const slide = requireSlide(op.slideIndex)
      if (!slide) return null
      await pushHistory()
      const toEmu = toEmuFor(op.fitWidthPx)
      for (const item of op.items ?? []) {
        const el = slide.elements.find((candidate) => candidate.id === item.sourceId)
        if (!el) continue
        el.transform = {
          ...el.transform,
          offset: {
            x: toEmu(item.xPx),
            y: toEmu(item.yPx),
            cx: toEmu(item.wPx),
            cy: toEmu(item.hPx),
          },
          rot: Math.round(item.rotationDeg * 60000),
        }
        el.dirtyTransform = true
      }
      updateConnectorsForMoved(
        slide,
        (op.items ?? []).map((item: any) => item.sourceId),
      )
      session.fitWidthPx = op.fitWidthPx
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    addElement: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      const slide = requireSlide(op.slideIndex)
      if (!slide) return null
      await pushHistory()
      const toEmu = toEmuFor(op.fitWidthPx)
      const paragraphs: Paragraph[] | undefined = op.paragraphs?.length
        ? (op.paragraphs as Paragraph[])
        : op.text
          ? String(op.text)
              .split('\n')
              .map((text) => ({ runs: [{ text }] }))
          : undefined
      const el = addElement(slide, {
        kind: op.kind,
        offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
        ...(paragraphs ? { paragraphs } : {}),
        ...(op.fillColor ? { fillColor: op.fillColor } : {}),
        ...(op.stroke
          ? { stroke: { color: op.stroke.color, widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT) } }
          : {}),
      })
      session.fitWidthPx = op.fitWidthPx
      setDirtyState(true)
      const slideRender = rebuild(op.slideIndex)
      return slideRender ? { slide: slideRender, sourceId: el.id } : null
    },
    deleteElement: async (op: any) => {
      if (mode !== 'edit') return null
      const slide = requireSlide(op.slideIndex)
      if (!slide?.elements.some((item) => item.id === op.sourceId)) return null
      await pushHistory()
      if (!deleteElement(slide, op.sourceId)) return null
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    editFill: async (op: any) => {
      if (mode !== 'edit' || op.groupId) return null
      const slide = requireSlide(op.slideIndex)
      const el = slide?.elements.find((item) => item.id === op.sourceId)
      if (!slide || !el) return null
      await pushHistory()
      if (typeof op.fill === 'string') {
        el.fill = op.fill === 'none' ? { type: 'none' } : { type: 'solid', color: op.fill }
      } else {
        const gradient = op.fill.gradient
        el.fill = {
          type: 'gradient',
          stops: [
            { pos: 0, color: gradient.from },
            { pos: 1, color: gradient.to },
          ],
          ...(gradient.radial
            ? { path: 'circle' as const }
            : { angle: Math.round((gradient.angleDeg ?? 0) * 60000) }),
        }
      }
      el.dirtyFill = true
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    editStroke: async (op: any) => {
      if (mode !== 'edit' || op.groupId) return null
      const slide = requireSlide(op.slideIndex)
      const el = slide?.elements.find((item) => item.id === op.sourceId)
      if (!slide || !el) return null
      await pushHistory()
      el.stroke = op.stroke
        ? {
            fill: { type: 'solid', color: op.stroke.color },
            width: Math.round(op.stroke.widthPt * EMU_PER_PT),
            ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
          }
        : undefined
      el.dirtyStroke = true
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    editBackground: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      const targets = op.allSlides ? session.opened.deck.slides : [requireSlide(op.slideIndex)].filter(Boolean)
      if (!targets.length) return null
      await pushHistory()
      for (const slide of targets as Slide[]) setSlideBackground(slide, op.color)
      session.fitWidthPx = op.fitWidthPx ?? session.fitWidthPx
      setDirtyState(true)
      return buildAllSlides(session.opened, session.fitWidthPx)
    },
    addBlankSlide: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      const slide = insertBlankSlide(session.opened, op.sourceIndex)
      if (!slide) return null
      session.fitWidthPx = op.fitWidthPx
      setDirtyState(true)
      return { slides: buildAllSlides(session.opened, op.fitWidthPx), index: op.sourceIndex + 1 }
    },
    addSlide: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      const slide = duplicateSlide(session.opened, op.sourceIndex, { clearText: Boolean(op.clearText) })
      if (!slide) return null
      session.fitWidthPx = op.fitWidthPx
      setDirtyState(true)
      return { slides: buildAllSlides(session.opened, op.fitWidthPx), index: op.sourceIndex + 1 }
    },
    addSlideWithLayout: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      let layoutPath: string | undefined = op.layoutPath
      if (layoutPath?.startsWith(BUILTIN_LAYOUT_PREFIX)) {
        layoutPath =
          ensureBuiltinLayout(
            session.opened.archive,
            session.opened.deck.size,
            layoutPath.slice(BUILTIN_LAYOUT_PREFIX.length),
          ) ?? undefined
      }
      const slide = layoutPath
        ? insertSlideWithLayout(session.opened, op.sourceIndex, layoutPath)
        : null
      if (!slide) return null
      session.fitWidthPx = op.fitWidthPx
      setDirtyState(true)
      return { slides: buildAllSlides(session.opened, op.fitWidthPx), index: op.sourceIndex + 1 }
    },
    deleteSlide: async (index: number) => {
      if (mode !== 'edit' || !session || session.opened.deck.slides.length <= 1) return null
      await pushHistory()
      if (!deleteSlide(session.opened, index)) return null
      setDirtyState(true)
      return buildAllSlides(session.opened, session.fitWidthPx)
    },
    moveSlide: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      if (!moveSlideModel(session.opened, op.fromIndex, op.toIndex)) return null
      setDirtyState(true)
      return {
        slides: buildAllSlides(session.opened, session.fitWidthPx),
        sections: getSections(session.opened),
      }
    },
    setSlideLayout: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      let layoutPath: string | undefined = op.layoutPath
      if (layoutPath?.startsWith(BUILTIN_LAYOUT_PREFIX)) {
        layoutPath =
          ensureBuiltinLayout(
            session.opened.archive,
            session.opened.deck.size,
            layoutPath.slice(BUILTIN_LAYOUT_PREFIX.length),
          ) ?? undefined
      }
      const result = layoutPath
        ? setSlideLayout(session.opened, op.slideIndex, layoutPath)
        : resetSlideLayout(session.opened, op.slideIndex)
      if (!result) return null
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    setSlideSize: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      if (!setSlideSize(session.opened, op.cx, op.cy)) return null
      setDirtyState(true)
      return buildAllSlides(session.opened, session.fitWidthPx)
    },
    insertImage: async (slideIndex: number, fitWidthPx: number) => {
      if (mode !== 'edit' || !session) return null
      const slide = requireSlide(slideIndex)
      if (!slide) return null
      const selected = await host.pickFile({
        multiple: false,
        accept: [...IMAGE_MIMES, '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg'],
        mode: 'file',
      })
      if (!selected?.[0]) return null
      const file = await selectedToOfficeFile(host, selected[0])
      const ext = extensionOf(file)
      if (!isImageMime(file.mimeType) && !['png', 'jpg', 'gif', 'bmp', 'webp', 'svg'].includes(ext)) {
        return { error: 'unsupported' as const, ext }
      }
      await pushHistory()
      const deck = session.opened.deck.size
      const cx = Math.round(deck.cx * 0.45)
      const cy = Math.round(deck.cy * 0.45)
      const el = addPicture(session.opened, slide, {
        bytes: new Uint8Array(file.bytes),
        ext,
        offset: {
          x: Math.round((deck.cx - cx) / 2),
          y: Math.round((deck.cy - cy) / 2),
          cx,
          cy,
        },
        name: file.name,
      })
      if (!el) return { error: 'unsupported' as const, ext }
      session.fitWidthPx = fitWidthPx
      setDirtyState(true)
      const rendered = rebuild(slideIndex)
      return rendered ? { slide: rendered, sourceId: el.id } : null
    },
    addImageBytes: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      const slide = requireSlide(op.slideIndex)
      if (!slide) return null
      await pushHistory()
      const toEmu = toEmuFor(op.fitWidthPx)
      const el = addPicture(session.opened, slide, {
        bytes: base64ToBytes(op.base64),
        ext: op.ext,
        offset: {
          x: toEmu(op.xPx),
          y: toEmu(op.yPx),
          cx: Math.max(1, toEmu(op.wPx)),
          cy: Math.max(1, toEmu(op.hPx)),
        },
        ...(op.name ? { name: op.name } : {}),
      })
      if (!el) return { error: 'unsupported' as const, ext: op.ext }
      session.fitWidthPx = op.fitWidthPx
      setDirtyState(true)
      const rendered = rebuild(op.slideIndex)
      return rendered ? { slide: rendered, sourceId: el.id } : null
    },
    getTransition: async (index: number) => {
      const slide = requireSlide(index)
      return slide ? getSlideTransition(slide) : 'none'
    },
    getAnimations: async () => [],
    getSections: async () => (session ? getSections(session.opened) : []),
    setSections: async (sections: any[]) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      setSections(session.opened, sections)
      setDirtyState(true)
      return getSections(session.opened)
    },
    addSection: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      const result = addSection(session.opened, op.atSlideIndex, op.name)
      if (result) setDirtyState(true)
      return result
    },
    renameSection: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      const result = renameSection(session.opened, op.id, op.name)
      if (result) setDirtyState(true)
      return result
    },
    removeSection: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      const result = removeSection(session.opened, op.id, { keepSlides: true })
      if (result) setDirtyState(true)
      return result
    },
    moveSection: async (op: any) => {
      if (mode !== 'edit' || !session) return null
      await pushHistory()
      const sections = moveSection(session.opened, op.id, op.dir)
      if (!sections) return null
      setDirtyState(true)
      return { slides: buildAllSlides(session.opened, session.fitWidthPx), sections }
    },
    getNotes: async (index: number) => {
      const slide = requireSlide(index)
      return session && slide ? getSlideNotes(session.opened.archive, slide.path) : ''
    },
    setNotes: async (op: any) => {
      if (mode !== 'edit' || !session) return false
      await pushHistory()
      const ok = setSlideNotes(session.opened, op.slideIndex, op.text)
      if (ok) setDirtyState(true)
      return ok
    },
    getComments: async (index: number) => {
      const slide = requireSlide(index)
      return session && slide ? getSlideComments(session.opened.archive, slide.path) : []
    },
    setSlideHidden: async (op: any) => {
      if (mode !== 'edit') return null
      const slide = requireSlide(op.slideIndex)
      if (!slide) return null
      await pushHistory()
      setSlideHidden(slide, op.hidden)
      setDirtyState(true)
      return rebuild(op.slideIndex)
    },
    undo: async () => {
      if (!session || session.undoStack.length === 0) return null
      const current = await snapshot()
      const previous = session.undoStack.pop()
      if (!previous) return null
      if (current) session.redoStack.push(current)
      return restoreSnapshot(previous)
    },
    redo: async () => {
      if (!session || session.redoStack.length === 0) return null
      const current = await snapshot()
      const next = session.redoStack.pop()
      if (!next) return null
      if (current) session.undoStack.push(current)
      return restoreSnapshot(next)
    },
    save: () => persist('save'),
    saveAs: (defaultName: string) => persist('saveAs', defaultName),
    saveHistoryVersion,
    exportPptx,
    requestHostClose: async () => host.requestClose?.(),
    isDirty: async () => dirty,
    getRecentFiles: async () => [],
    onMenuCommand: (handler: (command: MenuCommand) => void) => {
      menuHandlers.add(handler)
      return () => menuHandlers.delete(handler)
    },
    onOpened: (handler: (result: OpenResult) => void) => {
      openedHandlers.add(handler)
      return () => openedHandlers.delete(handler)
    },
    onRenamed: (handler: (path: string) => void) => {
      renamedHandlers.add(handler)
      return () => renamedHandlers.delete(handler)
    },
    onCloseSaveRequest: () => () => {},
    reportCloseSaveResult: () => {},
    setAutoSavePref: () => {},
    getAiSettings: async () => null,
    setAiSettings: async () => {},
    cloudGenStatus: async () => ({ enabled: false }),
    gskStatus: async () => ({ available: false }),
    aiGskStatus: async () => ({ loggedIn: false }),
    getShapeKeys: async () => [],
    getSlideLinks: async () => [],
    getRunLinks: async () => [],
    getHeaderFooter: async () => ({ footer: null, slideNum: false, date: null }),
    hasSlideClipboard: async () => false,
    clipboardExternal: async () => ({ kind: 'none' as const }),
    nativeClipboard: async (op: 'cut' | 'copy' | 'paste') => document.execCommand(op),
    beginHistoryBatch: async () => false,
    endHistoryBatch: async () => null,
    getChartColorSchemes: async () => null,
    getChartData: async () => null,
    getMediaData: async () => null,
    pickExportDir: async () => null,
    pickExportPdfPath: async () => null,
    exportImages: async () => ({ ok: false, error: 'Image export is unavailable in Web mode.' }),
    exportPdf: async () => ({ ok: false, error: 'PDF export is unavailable in Web mode.' }),
    printSlides: async () => ({ ok: false, error: 'Printing is unavailable in Web mode.' }),
    masterEnter: async () => null,
    masterOpen: async () => null,
    masterClose: async () => null,
    presenterStart: async () => ({ audience: false }),
    presenterSync: () => {},
    presenterInk: () => {},
    presenterSwap: async () => false,
    presenterEnd: async () => {},
    audienceReady: async () => null,
    audienceNav: () => {},
    onShowSync: () => () => {},
    onShowInk: () => () => {},
    onAudienceNav: () => () => {},
    onAiStream: () => () => {},
  }

  const fallback = (property: PropertyKey): unknown => {
    const name = String(property)
    if (name.startsWith('on')) return () => () => {}
    if (['setAutoSavePref', 'presenterSync', 'presenterInk', 'audienceNav'].includes(name)) {
      return () => {}
    }
    return async () => null
  }

  const slidesApi = new Proxy(explicit, {
    get(target, property) {
      return Reflect.has(target, property) ? Reflect.get(target, property) : fallback(property)
    },
  }) as unknown as SlidesApi

  const unsubscribeBridge = bridge?.subscribe((message: HostToEditorMessage) => {
    switch (message.type) {
      case 'office:init': {
        if (message.payload.kind !== 'pptx') return
        setMode(message.payload.mode)
        if (message.payload.locale) setLanguage(message.payload.locale)
        const file = message.payload.file
        void openOfficeFile(file, lastFitWidthPx).then((result) => {
          if (initialOpenResolve) {
            const resolve = initialOpenResolve
            initialOpenResolve = null
            resolve(result)
          } else {
            for (const handler of openedHandlers) handler(result)
          }
        })
        break
      }
      case 'office:new': {
        if (message.payload.kind !== 'pptx') return
        setMode(message.payload.mode)
        if (message.payload.locale) setLanguage(message.payload.locale)
        currentFile = null
        session = null
        setDirtyState(false)
        host.setTitle('Untitled Presentation')
        if (initialOpenResolve) {
          const resolve = initialOpenResolve
          initialOpenResolve = null
          resolve(null)
        } else {
          void (slidesApi.newBlank(lastFitWidthPx) as Promise<OpenResult>).then((result) => {
            for (const handler of openedHandlers) handler(result)
          })
        }
        break
      }
      case 'office:set-locale':
        setLanguage(message.payload.locale)
        break
      case 'office:set-mode':
        setMode(message.payload.mode)
        break
      case 'office:save': {
        pendingSaveRequestIds.add(message.requestId)
        for (const handler of menuHandlers) handler('save')
        if (menuHandlers.size === 0) reportPendingHostSave(false, 'Editor save handler is not ready.')
        break
      }
      case 'office:query-state':
        bridge.send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:state-result',
          requestId: message.requestId,
          payload: {
            ready: true,
            dirty,
            saving,
            mode,
            title: title(),
          },
        })
        break
      case 'office:error':
        console.error(`[office host:${message.payload.code}] ${message.payload.message}`)
        break
      default:
        break
    }
  })

  document.documentElement.classList.toggle('office-view-mode', mode === 'view')

  return {
    slidesApi,
    destroy: () => {
      unsubscribeBridge?.()
      initialOpenResolve?.(null)
      initialOpenResolve = null
      languageHandlers.clear()
      modeHandlers.clear()
      menuHandlers.clear()
      openedHandlers.clear()
      renamedHandlers.clear()
      pendingSaveRequestIds.clear()
      session = null
      currentFile = null
      document.documentElement.classList.remove('office-view-mode')
    },
  }
}
