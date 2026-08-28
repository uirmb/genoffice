import { Plugin, PluginKey } from '@tiptap/pm/state'
import {
  getProtectedInlineDataImage,
  isProtectedInlineDataImage,
} from '../markdown/inlineDataImages'

const DEFER_THRESHOLD = 512 * 1024
const DEFERRED_ATTR = 'data-md-deferred-image'
const PLACEHOLDER =
  'data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 width=%22720%22 height=%22180%22 viewBox=%220 0 720 180%22%3E%3Crect width=%22720%22 height=%22180%22 rx=%228%22 fill=%22%23f4f6f8%22/%3E%3Crect x=%220.5%22 y=%220.5%22 width=%22719%22 height=%22179%22 rx=%227.5%22 fill=%22none%22 stroke=%22%23d9dde2%22 stroke-dasharray=%226 6%22/%3E%3Cg fill=%22none%22 stroke=%22%23a8b0ba%22 stroke-width=%224%22 stroke-linecap=%22round%22 stroke-linejoin=%22round%22%3E%3Crect x=%22324%22 y=%2257%22 width=%2272%22 height=%2266%22 rx=%228%22/%3E%3Ccircle cx=%22348%22 cy=%2279%22 r=%228%22/%3E%3Cpath d=%22m333 112 18-18 14 14 10-10 16 14%22/%3E%3C/g%3E%3C/svg%3E'

const SESSION = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
const directDeferredSources = new Map<string, string>()
const queuedElements = new WeakSet<HTMLImageElement>()
const loadQueue: Array<{ element: HTMLImageElement; key: string }> = []
let sequence = 0
let activeLoad = false
let queueStartScheduled = false

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

  return {
    src: PLACEHOLDER,
    [DEFERRED_ATTR]: key,
    'aria-busy': 'true',
    decoding: 'async',
    loading: 'lazy',
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

    let settled = false
    const watchdog = setTimeout(() => {
      if (settled) return
      settled = true
      finishLoad(element, key, source, true)
    }, 30_000)

    const settle = (success: boolean) => {
      if (settled) return
      settled = true
      clearTimeout(watchdog)
      finishLoad(element, key, source, success)
    }

    loader.onload = () => settle(true)
    loader.onerror = () => settle(false)
    loader.src = source
    return
  }
}

function scheduleQueueStart(): void {
  if (activeLoad || queueStartScheduled) return
  queueStartScheduled = true

  const start = () => {
    queueStartScheduled = false
    runNext()
  }

  // Two frames guarantee that text + SVG placeholders get a paint opportunity
  // before any multi-megabyte Base64 image is handed to the decoder.
  if (typeof requestAnimationFrame === 'function') {
    requestAnimationFrame(() => requestAnimationFrame(start))
  } else {
    setTimeout(start, 0)
  }
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
