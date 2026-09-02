import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { buildPrintHtml } from '../src/renderer/export/printHtml'

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

function firstImageSrc(editor: Editor): string {
  return String(editor.getJSON().content?.find((node) => node.type === 'image')?.attrs?.src ?? '')
}

function largePngDataUrl(width: number, height: number): string {
  const header = new Uint8Array(24)
  header.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0)
  header.set([0x49, 0x48, 0x44, 0x52], 12)
  header[16] = (width >>> 24) & 0xff
  header[17] = (width >>> 16) & 0xff
  header[18] = (width >>> 8) & 0xff
  header[19] = width & 0xff
  header[20] = (height >>> 24) & 0xff
  header[21] = (height >>> 16) & 0xff
  header[22] = (height >>> 8) & 0xff
  header[23] = height & 0xff

  let binary = ''
  for (const byte of header) binary += String.fromCharCode(byte)
  return `data:image/png;base64,${btoa(binary)}${'A'.repeat(1024 * 1024)}`
}

describe('deferred inline image display', () => {
  it('renders a large Base64 image as a same-size lightweight placeholder while keeping the authored src', () => {
    const editor = createEditor()
    const dataUrl = largePngDataUrl(1920, 1080)

    expect(
      editor.commands.setContent({
        type: 'doc',
        content: [{ type: 'image', attrs: { src: dataUrl, alt: 'large' } }],
      }),
    ).toBe(true)

    const image = editor.view.dom.querySelector('img')
    expect(image).not.toBeNull()
    expect(image?.getAttribute('src')).toMatch(/^data:image\/svg\+xml,/)
    expect(image?.getAttribute('width')).toBe('1920')
    expect(image?.getAttribute('height')).toBe('1080')
    expect(image?.getAttribute('data-md-deferred-image')).toBeTruthy()
    expect(image?.getAttribute('aria-busy')).toBe('true')
    expect(image?.getAttribute('loading')).toBe('eager')
    expect(firstImageSrc(editor)).toBe(dataUrl)
    expect(editor.getMarkdown()).toContain(dataUrl)
  })

  it('keeps small inline images on the normal immediate display path', () => {
    const editor = createEditor()
    const dataUrl = `data:image/png;base64,${'A'.repeat(1024)}`

    editor.commands.setContent({
      type: 'doc',
      content: [{ type: 'image', attrs: { src: dataUrl, alt: 'small' } }],
    })

    const image = editor.view.dom.querySelector('img')
    expect(image?.getAttribute('src')).toBe(dataUrl)
    expect(image?.hasAttribute('data-md-deferred-image')).toBe(false)
    expect(image?.getAttribute('decoding')).toBe('async')
    expect(image?.getAttribute('loading')).toBe('lazy')
  })

  it('materializes the real Base64 image in PDF print HTML even before screen loading finishes', () => {
    const editor = createEditor()
    const dataUrl = `data:image/png;base64,${'B'.repeat(1024 * 1024)}`

    editor.commands.setContent({
      type: 'doc',
      content: [
        { type: 'paragraph', content: [{ type: 'text', text: 'Before image' }] },
        { type: 'image', attrs: { src: dataUrl, alt: 'export-large' } },
      ],
    })

    const html = buildPrintHtml(editor.view.dom, 'Deferred image export')
    expect(html).toContain(dataUrl)
    expect(html).not.toContain('data-md-deferred-image')
  })
})
