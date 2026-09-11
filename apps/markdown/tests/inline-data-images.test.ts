import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  getProtectedInlineDataImage,
  isProtectedInlineDataImage,
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

function findImageSrcs(node: unknown, result: string[] = []): string[] {
  if (!node || typeof node !== 'object') return result
  const value = node as {
    type?: string
    attrs?: Record<string, unknown>
    content?: unknown[]
  }
  if (value.type === 'image') result.push(String(value.attrs?.src ?? ''))
  for (const child of value.content ?? []) findImageSrcs(child, result)
  return result
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

  it('keeps the first synchronous setContent lightweight while serialization stays lossless', () => {
    const editor = createEditor()
    const payload = 'A'.repeat(20 * 1024 * 1024)
    const dataUrl = `data:image/jpeg;base64,${payload}`
    const markdown = `# Large image\n\n![photo](${dataUrl})\n\nAfter image\n`

    expect(editor.commands.setContent(markdown, { contentType: 'markdown' })).toBe(true)

    const [storedSrc] = findImageSrcs(editor.getJSON())
    expect(storedSrc).toBeTruthy()
    expect(isProtectedInlineDataImage(storedSrc!)).toBe(true)
    expect(storedSrc!.length).toBeLessThan(128)
    expect(getProtectedInlineDataImage(storedSrc!)).toBe(dataUrl)
    expect(JSON.stringify(editor.getJSON()).length).toBeLessThan(2048)
    expect(editor.view.dom.textContent).toContain('Large image')
    expect(editor.view.dom.textContent).toContain('After image')

    const image = editor.view.dom.querySelector('img')
    expect(image?.getAttribute('src')).toMatch(/^data:image\/svg\+xml,/)
    expect(image?.getAttribute('data-md-deferred-image')).toBe(storedSrc)

    const serialized = editor.getMarkdown()
    expect(serialized).toContain(dataUrl)
    expect(serialized).toContain('After image')
    expect(serialized).not.toContain('genoffice.invalid/__inline-data-image/')
  })

  it('loads a 46 MB Markdown body into a small first-paint document with four placeholders', () => {
    const editor = createEditor()
    const sizes = [12, 12, 11, 11].map((mb) => mb * 1024 * 1024)
    const dataUrls = sizes.map(
      (size, index) => `data:image/png;base64,${String.fromCharCode(65 + index).repeat(size)}`,
    )
    const markdown = [
      '# Large multi-image document',
      ...dataUrls.flatMap((dataUrl, index) => [
        `## Image ${index + 1}`,
        `![image-${index + 1}](${dataUrl})`,
        `Text after image ${index + 1}.`,
      ]),
      '',
    ].join('\n\n')

    expect(markdown.length).toBeGreaterThan(46 * 1024 * 1024)
    expect(editor.commands.setContent(markdown, { contentType: 'markdown' })).toBe(true)

    const storedSrcs = findImageSrcs(editor.getJSON())
    expect(storedSrcs).toHaveLength(4)
    expect(storedSrcs.every(isProtectedInlineDataImage)).toBe(true)
    expect(storedSrcs.map((src) => getProtectedInlineDataImage(src))).toEqual(dataUrls)
    expect(JSON.stringify(editor.getJSON()).length).toBeLessThan(16 * 1024)
    expect(editor.view.dom.querySelectorAll('img[data-md-deferred-image]')).toHaveLength(4)
    expect(editor.view.dom.textContent).toContain('Text after image 4.')

    const serialized = editor.getMarkdown()
    for (const dataUrl of dataUrls) expect(serialized).toContain(dataUrl)
  })

  it('does not rewrite a data URL that is not a Markdown image destination', () => {
    const markdown = '`data:image/png;base64,AAAA`\n\n[data](data:image/png;base64,BBBB)'
    expect(protectInlineDataImages(markdown)).toBe(markdown)
  })
})
