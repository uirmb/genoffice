import { afterEach, describe, expect, it, vi } from 'vitest'

import type { WorkbookFile } from '../src/shared/desktop-api'
import { readXlsxWorkbookMedia } from '../src/web/engine-client'

const sessionId = '00000000-0000-4000-8000-000000000001'
const mediaPath = 'xl/media/image1.png'

function workbook(): WorkbookFile {
  return {
    sessionId,
    name: 'media.xlsx',
    sha256: '0'.repeat(64),
    entryCount: 1,
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet1',
        rowCount: 1,
        columnCount: 1,
        columnWidths: [],
        defaultRowHeight: null,
        defaultColumnWidth: null,
        freeze: null,
        hidden: false,
        tabColor: null,
        showGridLines: true,
        showFormulas: false,
        tables: [],
        comments: [],
        pivotRanges: [],
        pivotTables: [],
        sparklines: [],
      },
    ],
    styles: [],
    dxfStyles: [],
    visuals: [
      {
        id: 'image-1',
        sheetId: 'sheet-1',
        kind: 'image',
        anchor: {
          fromRow: 0,
          fromColumn: 0,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 1,
          toColumn: 1,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
        mediaPath,
        mediaType: 'image/png',
      },
    ],
    definedNames: [],
    readOnly: false,
  }
}

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('Sheets Web XLSX media client', () => {
  it('reads image bytes from the workbook session archive', async () => {
    const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const fetchMock = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            entries: [
              {
                name: mediaPath,
                crc32: 0,
                compressedSize: bytes.length,
                uncompressedSize: bytes.length,
              },
            ],
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        ),
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            entries: [
              {
                name: mediaPath,
                contentBase64: Buffer.from(bytes).toString('base64'),
              },
            ],
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        ),
      )
    vi.stubGlobal('fetch', fetchMock)

    const result = await readXlsxWorkbookMedia(
      { sessionId, visualId: 'image-1' },
      workbook(),
    )

    expect(result).toEqual({
      mediaType: 'image/png',
      base64: Buffer.from(bytes).toString('base64'),
    })
    expect(fetchMock).toHaveBeenCalledTimes(2)
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain('/archive/manifest')
    expect(String(fetchMock.mock.calls[1]?.[0])).toContain('/archive/read')
  })

  it('rejects oversized images before downloading their bytes', async () => {
    const fetchMock = vi.fn<typeof fetch>().mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          entries: [
            {
              name: mediaPath,
              crc32: 0,
              compressedSize: 1,
              uncompressedSize: 20 * 1024 * 1024 + 1,
            },
          ],
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      ),
    )
    vi.stubGlobal('fetch', fetchMock)

    await expect(
      readXlsxWorkbookMedia({ sessionId, visualId: 'image-1' }, workbook()),
    ).rejects.toThrow('20MB preview limit')
    expect(fetchMock).toHaveBeenCalledTimes(1)
  })
})
