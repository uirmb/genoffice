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

describe('deferred inline image display', () => {
  it('renders a large Base64 image as a lightweight placeholder while keeping the authored src', () => {
    const editor = createEditor()
    const dataUrl = `data:image/jpeg;base64,${'A'.repeat(1024 * 1024)}`

    expect(
      editor.commands.setContent({
        type: 'doc',
        content: [{ type: 'image', attrs: { src: dataUrl, alt: 'large' } }],
      }),
    ).toBe(true)

    const image = editor.view.dom.querySelector('img')
    expect(image).not.toBeNull()
    expect(image?.getAttribute('src')).toMatch(/^data:image\/svg\+xml,/)
    expect(image?.getAttribute('data-md-deferred-image')).toBeTruthy()
    expect(image?.getAttribute('aria-busy')).toBe('true')
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
