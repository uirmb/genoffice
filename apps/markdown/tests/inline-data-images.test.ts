import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  protectInlineDataImages,
  restoreInlineDataImage,
} from '../src/renderer/markdown/inlineDataImages'

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

function findImageSrc(node: unknown): string | null {
  if (!node || typeof node !== 'object') return null
  const value = node as {
    type?: string
    attrs?: Record<string, unknown>
    content?: unknown[]
  }
  if (value.type === 'image') return String(value.attrs?.src ?? '')
  for (const child of value.content ?? []) {
    const src = findImageSrc(child)
    if (src !== null) return src
  }
  return null
}

describe('inline data image protection', () => {
  it('replaces only the Base64 payload with a short lexer-safe URL and restores it losslessly', () => {
    const dataUrl = `data:image/png;base64,${'A'.repeat(1024 * 1024)}`
    const markdown = `before\n\n![large](${dataUrl})\n\nafter`

    const protectedMarkdown = protectInlineDataImages(markdown)
    expect(protectedMarkdown.length).toBeLessThan(256)
    expect(protectedMarkdown).toContain('https://genoffice.invalid/__inline-data-image/')

    const placeholder = protectedMarkdown.match(/\((https:\/\/genoffice\.invalid\/[^)]+)\)/)?.[1]
    expect(placeholder).toBeTruthy()
    expect(restoreInlineDataImage(placeholder!)).toBe(dataUrl)
  })

  it('parses and serializes a 20 MB embedded image without exposing the Base64 body to MarkedJS', async () => {
    const editor = createEditor()
    // Tiptap fires extension create hooks after construction; yield once before parsing.
    await new Promise<void>((resolve) => setTimeout(resolve, 0))

    const payload = 'A'.repeat(20 * 1024 * 1024)
    const dataUrl = `data:image/jpeg;base64,${payload}`
    const markdown = `# Large image\n\n![photo](${dataUrl})\n\nAfter image\n`

    const json = editor.markdown.parse(markdown)
    expect(findImageSrc(json)).toBe(dataUrl)

    const serialized = editor.markdown.serialize(json)
    expect(serialized).toContain(dataUrl)
    expect(serialized).toContain('After image')
  })

  it('does not rewrite a data URL that is not a Markdown image destination', () => {
    const markdown = '`data:image/png;base64,AAAA`\n\n[data](data:image/png;base64,BBBB)'
    expect(protectInlineDataImages(markdown)).toBe(markdown)
  })
})
