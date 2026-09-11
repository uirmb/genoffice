import { Plugin, PluginKey } from '@tiptap/pm/state'
import {
  getProtectedInlineDataImage,
  isProtectedInlineDataImage,
} from '../markdown/inlineDataImages'
import { readInlineImageDimensions, type InlineImageDimensions } from './inlineImageDimensions'

const DEFER_THRESHOLD = 512 * 1024
const DEFERRED_ATTR = 'data-md-deferred-image'
const FALLBACK_DIMENSIONS: InlineImageDimensions = { width: 720, height: 180 }

const SESSION = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
const directDeferredSources = new Map<string, string>()
const queuedElements = new WeakSet<HTMLImageElement>()
const loadQueue: Array<{ element: HTMLImageElement; key: string }> = []
let sequence = 0
let activeLoad = false
let queueStartScheduled = false
let visiblePaintLogged = false

function now(): number {
  return typeof performance !== 'undefined' ? performance.now() : Date.now()
}

function perf(stage: string, extra: Record<string, unknown> = {}): void {
  console.info('[markdown:perf]', { stage, t: Math.round(now()), ...extra })
}

function placeholderDataUrl(dimensions: InlineImageDimensions | null): string {
  const { width, height } = dimensions ?? FALLBACK_DIMENSIONS
  const minDimension = Math.min(width, height)
  const strokeWidth = Math.max(1, minDimension * 0.004)
  const radius = Math.max(4, minDimension * 0.02)
  const iconSize = Math.max(48, Math.min(220, minDimension * 0.16))
  const iconWidth = iconSize
  const iconHeight = iconSize * 0.88
  const iconX = (width - iconWidth) / 2
  const iconY = (height - iconHeight) / 2
  const iconStroke = Math.max(2, iconSize * 0.045)

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="${width}" height="${height}" rx="${radius}" fill="#f4f6f8"/><rect x="${strokeWidth / 2}" y="${strokeWidth / 2}" width="${Math.max(1, width - strokeWidth)}" height="${Math.max(1, height - strokeWidth)}" rx="${radius}" fill="none" stroke="#d9dde2" stroke-width="${strokeWidth}" stroke-dasharray="${strokeWidth * 5} ${strokeWidth * 5}"/><g transform="translate(${iconX} ${iconY})" fill="none" stroke="#a8b0ba" stroke-width="${iconStroke}" stroke-linecap="round" stroke-linejoin="round"><rect x="0" y="0" width="${iconWidth}" height="${iconHeight}" rx="${iconSize * 0.1}"/><circle cx="${iconSize * 0.33}" cy="${iconSize * 0.3}" r="${iconSize * 0.11}"/><path d="M ${iconSize * 0.13} ${iconSize * 0.73} L ${iconSize * 0.38} ${iconSize * 0.48} L ${iconSize * 0.57} ${iconSize * 0.66} L ${iconSize * 0.72} ${iconSize * 0.52} L ${iconSize * 0.9} ${iconSize * 0.69}"/></g></svg>`
  return `data:image/svg+xml,${encodeURIComponent(svg)}`
}

export function shouldDeferInlineImage(src: string): boolean {
  return (
    isProtectedInlineDataImage(src) ||
    (src.length >= DEFER_THRESHOLD &&
      src.startsWith('data:image/') &&
      src.slice(0, 128).includes(';base64,'))
  )
}

function sourceForKey(key: string): string | null {
  return getProtectedInlineDataImage(key) ?? directDeferredSources.get(key) ?? null
}

export function deferredInlineImageAttributes(src: string): Record<string, string> | null {
  if (!shouldDeferInlineImage(src)) return null

  let key = src
  if (!isProtectedInlineDataImage(src)) {
    key = `${SESSION}-${sequence}`
    sequence += 1
    directDeferredSources.set(key, src)
  }

  const source = sourceForKey(key)
  const dimensions = source ? readInlineImageDimensions(source) : null

  return {
    src: placeholderDataUrl(dimensions),
    ...(dimensions
      ? {
          width: String(dimensions.width),
          height: String(dimensions.height),
        }
      : {}),
    [DEFERRED_ATTR]: key,
    'aria-busy': 'true',
    decoding: 'async',
    // The placeholder is tiny and must paint immediately. The large authored
    // image is still gated by the explicit decode queue below.
    loading: 'eager',
  }
}

function finishLoad(
  element: HTMLImageElement,
  key: string,
  source: string,
  success: boolean,
): void {
  if (success && element.isConnected) {
    // Only the live DOM is materialized. The editor node deliberately keeps the
    // short protected URL so the ProseMirror document remains lightweight.
    element.src = source
    element.removeAttribute(DEFERRED_ATTR)
    element.removeAttribute('aria-busy')
    directDeferredSources.delete(key)
  } else if (element.isConnected) {
    element.setAttribute('aria-busy', 'false')
  } else {
    directDeferredSources.delete(key)
  }

  queuedElements.delete(element)
  activeLoad = false
  setTimeout(runNext, 0)
}

function runNext(): void {
  if (activeLoad) return

  while (loadQueue.length > 0) {
    const item = loadQueue.shift()!
    const { element, key } = item
    const source = sourceForKey(key)

    if (!source || !element.isConnected) {
      queuedElements.delete(element)
      if (!element.isConnected) directDeferredSources.delete(key)
      continue
    }

    const ImageCtor = element.ownerDocument.defaultView?.Image
    if (!ImageCtor) {
      finishLoad(element, key, source, true)
      return
    }

    activeLoad = true
    const loader = new ImageCtor()
    loader.decoding = 'async'
    const startedAt = now()
    perf('image-decode-start', { chars: source.length })

    let settled = false
    const watchdog = setTimeout(() => {
      if (settled) return
      settled = true
      perf('image-decode-finish', {
        ms: Math.round(now() - startedAt),
        success: true,
        watchdog: true,
      })
      finishLoad(element, key, source, true)
    }, 30_000)

    const settle = (success: boolean) => {
      if (settled) return
      settled = true
      clearTimeout(watchdog)
      perf('image-decode-finish', { ms: Math.round(now() - startedAt), success })
      finishLoad(element, key, source, success)
    }

    loader.onload = () => settle(true)
    loader.onerror = () => settle(false)
    loader.src = source
    return
  }
}

function elementCanPaint(element: HTMLImageElement): boolean {
  return element.isConnected && element.getClientRects().length > 0
}

function hasPaintableQueuedImage(): boolean {
  return loadQueue.some(
    ({ element, key }) => Boolean(sourceForKey(key)) && elementCanPaint(element),
  )
}

function scheduleQueueStart(): void {
  if (activeLoad || queueStartScheduled) return
  queueStartScheduled = true

  const release = () => {
    queueStartScheduled = false
    runNext()
  }

  if (typeof requestAnimationFrame !== 'function') {
    setTimeout(release, 0)
    return
  }

  const waitUntilVisible = () => {
    if (loadQueue.length === 0) {
      queueStartScheduled = false
      return
    }

    // App.tsx historically hides .app-main while status=loading. Do not let the
    // decoder consume the large Base64 sources while those placeholders are
    // invisible; otherwise the user only sees the final image after loading.
    if (!hasPaintableQueuedImage()) {
      requestAnimationFrame(waitUntilVisible)
      return
    }

    if (!visiblePaintLogged) {
      visiblePaintLogged = true
      perf('placeholder-visible', { queued: loadQueue.length })
    }

    // The first frame observes that the editor is actually visible. Two more
    // frames guarantee at least one browser paint containing text + placeholders
    // before the first multi-megabyte data URL is handed to the image decoder.
    requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(release, 0)))
  }

  requestAnimationFrame(waitUntilVisible)
}

function enqueueDeferredImages(root: ParentNode): void {
  for (const element of root.querySelectorAll<HTMLImageElement>(`img[${DEFERRED_ATTR}]`)) {
    if (queuedElements.has(element)) continue
    const key = element.getAttribute(DEFERRED_ATTR)
    if (!key || !sourceForKey(key)) continue

    queuedElements.add(element)
    loadQueue.push({ element, key })
  }

  if (loadQueue.length > 0) scheduleQueueStart()
}

export function createDeferredInlineImagePlugin(): Plugin {
  return new Plugin({
    key: new PluginKey('deferredInlineImageLoader'),
    view(view) {
      enqueueDeferredImages(view.dom)
      return {
        update(nextView) {
          enqueueDeferredImages(nextView.dom)
        },
      }
    },
  })
}

/**
 * Print/export clones are not part of the live loading queue. Replace their
 * placeholder URLs with the authored Base64 source synchronously so exporting
 * before a large image has appeared on screen still produces the real image.
 */
export function materializeDeferredInlineImages(root: ParentNode): void {
  for (const element of root.querySelectorAll<HTMLImageElement>(`img[${DEFERRED_ATTR}]`)) {
    const key = element.getAttribute(DEFERRED_ATTR)
    const source = key ? sourceForKey(key) : null
    if (!source) continue

    element.src = source
    element.removeAttribute(DEFERRED_ATTR)
    element.removeAttribute('aria-busy')
  }
}
