import { Extension, type JSONContent } from '@tiptap/core'

const PROTECT_THRESHOLD = 512 * 1024
const PLACEHOLDER_SESSION = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
const PLACEHOLDER_ROOT = `https://genoffice.invalid/__inline-data-image/${PLACEHOLDER_SESSION}/`

const protectedImages = new Map<string, string>()
let sequence = 0

function nextImageDestination(
  markdown: string,
  from: number,
): { marker: number; dataStart: number; angleWrapped: boolean } | null {
  let searchFrom = from

  while (searchFrom < markdown.length) {
    const direct = markdown.indexOf('](data:image/', searchFrom)
    const angle = markdown.indexOf('](<data:image/', searchFrom)

    let marker = -1
    let dataStart = -1
    let angleWrapped = false
    if (direct >= 0 && (angle < 0 || direct < angle)) {
      marker = direct
      dataStart = direct + 2
    } else if (angle >= 0) {
      marker = angle
      dataStart = angle + 3
      angleWrapped = true
    }
    if (marker < 0) return null

    const imageOpen = markdown.lastIndexOf('![', marker)
    if (imageOpen >= 0 && markdown.indexOf(']', imageOpen + 2) === marker) {
      return { marker, dataStart, angleWrapped }
    }

    searchFrom = marker + 2
  }

  return null
}

/**
 * MarkedJS does not need to inspect the Base64 body of a large embedded image.
 * Very large data URLs can otherwise turn one Markdown line into tens of MB and
 * overflow the lexer stack. Small inline images stay on the normal immediate
 * path; only large Markdown image destinations are replaced before lexing.
 *
 * Base64 cannot contain ')' or '>', so the destination end is found with native
 * indexOf rather than tens of millions of JavaScript character checks. This is
 * important for first paint on 40+ MB self-contained Markdown files.
 *
 * The placeholder intentionally stays in the editor document. Keeping the tens
 * of megabytes of Base64 outside ProseMirror makes the first setContent cheap
 * enough to paint the text and image placeholders immediately. Serialization,
 * export and image display resolve the authored Base64 through this side table.
 */
export function protectInlineDataImages(markdown: string): string {
  protectedImages.clear()
  sequence = 0

  let cursor = 0
  let searchFrom = 0
  let output = ''

  while (searchFrom < markdown.length) {
    const destination = nextImageDestination(markdown, searchFrom)
    if (!destination) break

    const { marker, dataStart, angleWrapped } = destination
    const base64Marker = markdown.indexOf(';base64,', dataStart)
    if (base64Marker < 0 || base64Marker - dataStart > 96) {
      searchFrom = marker + 2
      continue
    }

    const mimePrefix = markdown.slice(dataStart, base64Marker)
    if (!/^data:image\/[a-z0-9.+-]+$/i.test(mimePrefix)) {
      searchFrom = marker + 2
      continue
    }

    const payloadStart = base64Marker + ';base64,'.length
    const payloadEnd = markdown.indexOf(angleWrapped ? '>' : ')', payloadStart)
    if (payloadEnd <= payloadStart) {
      searchFrom = marker + 2
      continue
    }

    if (payloadEnd - dataStart < PROTECT_THRESHOLD) {
      searchFrom = payloadEnd + 1
      continue
    }

    const placeholder = `${PLACEHOLDER_ROOT}${sequence}`
    sequence += 1
    protectedImages.set(placeholder, markdown.slice(dataStart, payloadEnd))

    output += markdown.slice(cursor, dataStart)
    output += placeholder
    cursor = payloadEnd
    searchFrom = payloadEnd + 1
  }

  if (cursor === 0) return markdown
  output += markdown.slice(cursor)
  return output
}

export function isProtectedInlineDataImage(src: string): boolean {
  return protectedImages.has(src)
}

export function getProtectedInlineDataImage(src: string): string | null {
  return protectedImages.get(src) ?? null
}

export function restoreInlineDataImage(src: string): string {
  return protectedImages.get(src) ?? src
}

export function restoreProtectedInlineDataImagesInMarkdown(markdown: string): string {
  let restored = markdown
  for (const [placeholder, source] of protectedImages) {
    if (restored.includes(placeholder)) restored = restored.split(placeholder).join(source)
  }
  return restored
}

export function restoreProtectedInlineDataImagesInJson(content: JSONContent): JSONContent {
  if (content.type === 'image' && typeof content.attrs?.src === 'string') {
    content.attrs.src = restoreInlineDataImage(content.attrs.src)
  }
  for (const child of content.content ?? []) restoreProtectedInlineDataImagesInJson(child)
  return content
}

/**
 * Tiptap's Markdown extension creates editor.markdown in its onBeforeCreate hook.
 * This protection extension has a lower priority, so its onBeforeCreate runs after
 * Markdown has installed the manager but before React effects can load a document.
 *
 * Parsing keeps protected placeholders in the editor document so ProseMirror does
 * not carry/copy huge Base64 strings during first render. Serialization restores
 * the exact authored Base64, preserving Save / Save As / history / download.
 */
export const InlineDataImageProtection = Extension.create({
  name: 'inlineDataImageProtection',
  priority: 50,
  onBeforeCreate() {
    const markdown = this.editor.markdown
    if (!markdown) return

    const originalParse = markdown.parse.bind(markdown)
    const originalSerialize = markdown.serialize.bind(markdown)

    markdown.parse = (source: string) => originalParse(protectInlineDataImages(source))
    markdown.serialize = (content: JSONContent) =>
      restoreProtectedInlineDataImagesInMarkdown(originalSerialize(content))
  },
})
