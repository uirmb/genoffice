import { describe, expect, it } from 'vitest'
import {
  buildBlankDocx,
  parseDocx,
  saveDocx,
  type GeneratedBlock,
  type SaveBlock,
} from '@genoffice/docx-engine'

function visibleBlocks<T extends { hidden?: boolean }>(blocks: T[]): T[] {
  return blocks.filter((block) => !block.hidden)
}

describe('Docs Web DOCX round-trip', () => {
  it('parses, edits, saves, and reparses a real DOCX package', async () => {
    const original = await buildBlankDocx({ eastAsiaFont: 'Microsoft YaHei' })
    const parsed = await parseDocx(original)
    const visible = visibleBlocks(parsed.blocks)

    expect(visible.length).toBeGreaterThan(0)
    const first = visible[0]
    expect(first.docxIndex).not.toBeNull()

    const editedText = 'GenOffice Web MVP round-trip 中文验证'
    const generated: GeneratedBlock = {
      type: first.type === 'heading' || first.type === 'listItem' ? first.type : 'paragraph',
      level: first.level,
      styleId: first.styleId,
      list: first.list,
      format: first.format,
      rawPPr: first.rawPPr,
      bookmarks: first.bookmarks,
      hiddenBookmarks: first.hiddenBookmarks,
      commentStarts: first.commentStarts,
      commentEnds: first.commentEnds,
      runs: [{ text: editedText }],
      sdtShell: first.sdtShell,
    }

    const finalBlocks: SaveBlock[] = visible.map((block, index) => {
      if (index === 0) return { kind: 'generated', block: generated }
      if (block.docxIndex === null) throw new Error('Expected an original DOCX block')
      return { kind: 'original', docxIndex: block.docxIndex }
    })

    const saved = await saveDocx(parsed, finalBlocks, {
      savedAt: '2026-08-08T00:00:00.000Z',
    })

    expect(saved.byteLength).toBeGreaterThan(0)
    expect(saved).not.toEqual(original)

    const reparsed = await parseDocx(saved)
    const text = visibleBlocks(reparsed.blocks)
      .flatMap((block) => block.runs ?? [])
      .map((run) => run.text)
      .join('')

    expect(text).toContain(editedText)
    expect(reparsed.internal.documentXml).toContain('GenOffice Web MVP round-trip')
    expect(reparsed.internal.documentXml).toContain('中文验证')
  })

  it('keeps an untouched DOCX byte-identical through the save engine', async () => {
    const original = await buildBlankDocx()
    const parsed = await parseDocx(original)
    const finalBlocks: SaveBlock[] = visibleBlocks(parsed.blocks).map((block) => {
      if (block.docxIndex === null) throw new Error('Expected an original DOCX block')
      return { kind: 'original', docxIndex: block.docxIndex }
    })

    const saved = await saveDocx(parsed, finalBlocks)
    expect(saved).toEqual(original)
  })
})
