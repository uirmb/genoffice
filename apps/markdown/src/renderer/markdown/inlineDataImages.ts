import { Extension, type JSONContent } from '@tiptap/core'

const PLACEHOLDER_SESSION = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
const PLACEHOLDER_ROOT = `https://genoffice.invalid/__inline-data-image/${PLACEHOLDER_SESSION}/`

const protectedImages = new Map<string, string>()
let sequence = 0

function isBase64Char(code: number): boolean {
  return (
    (code >= 48 && code <= 57) ||
    (code >= 65 && code <= 90) ||
    (code >= 97 && code <= 122) ||
    code === 43 ||
    code === 47 ||
    code === 45 ||
    code === 95 ||
    code === 61
  )
}

function nextImageDestination(
  markdown: string,
  from: number,
): { marker: number; dataStart: number } | null {
  let searchFrom = from

  while (searchFrom < markdown.length) {
    const direct = markdown.indexOf('](data:image/', searchFrom)
    const angle = markdown.indexOf('](<data:image/', searchFrom)

    let marker = -1
    let dataStart = -1
    if (direct >= 0 && (angle < 0 || direct < angle)) {
      marker = direct
      dataStart = direct + 2
    } else if (angle >= 0) {
      marker = angle
      dataStart = angle + 3
    }
    if (marker < 0) return null

    const imageOpen = markdown.lastIndexOf('![', marker)
    if (imageOpen >= 0 && markdown.indexOf(']', imageOpen + 2) === marker) {
      return { marker, dataStart }
    }

    searchFrom = marker + 2
  }

  return null
}

/**
 * MarkedJS does not need to inspect the Base64 body of an embedded image. Very
 * large data URLs can otherwise turn one Markdown line into tens of megabytes
 * and overflow the lexer stack. Replace only Markdown image destinations with a
 * short HTTPS placeholder before lexing.
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

    const { marker, dataStart } = destination
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
    let payloadEnd = payloadStart
    while (payloadEnd < markdown.length && isBase64Char(markdown.charCodeAt(payloadEnd))) {
      payloadEnd += 1
    }
    if (payloadEnd === payloadStart) {
      searchFrom = marker + 2
      continue
    }

    const placeholder = `${PLACEHOLDER_ROOT}${sequence}`
    sequence += 1
    protectedImages.set(placeholder, markdown.slice(dataStart, payloadEnd))

    output += markdown.slice(cursor, dataStart)
    output += placeholder
    cursor = payloadEnd
    searchFrom = payloadEnd
  }

  if (cursor === 0) return markdown
  output += markdown.slice(cursor)
  return output
}

export function restoreInlineDataImage(src: string): string {
  const original = protectedImages.get(src)
  if (!original) return src
  protectedImages.delete(src)
  return original
}

function restoreInlineDataImages(content: JSONContent): JSONContent {
  if (content.type === 'image' && typeof content.attrs?.src === 'string') {
    content.attrs.src = restoreInlineDataImage(content.attrs.src)
  }
  for (const child of content.content ?? []) restoreInlineDataImages(child)
  return content
}

/**
 * Tiptap's Markdown extension creates editor.markdown in its onBeforeCreate hook.
 * This protection extension has a lower priority, so its onBeforeCreate runs after
 * Markdown has installed the manager but before React effects can load a document.
 * Wrapping here is important: onCreate is too late for the first setContent call in
 * the Web editor and can let a large Base64 line reach MarkedJS unprotected.
 */
export const InlineDataImageProtection = Extension.create({
  name: 'inlineDataImageProtection',
  priority: 50,
  onBeforeCreate() {
    const markdown = this.editor.markdown
    if (!markdown) return
    const originalParse = markdown.parse.bind(markdown)
    markdown.parse = (source: string) =>
      restoreInlineDataImages(originalParse(protectInlineDataImages(source)))
  },
})
