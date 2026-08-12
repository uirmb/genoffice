import type { OfficeHostApi } from '@genoffice/office-host-api'
import { afterEach, describe, expect, it, vi } from 'vitest'

import type { WorkbookFile } from '../src/shared/desktop-api'
import { readXlsxWorkbookMedia } from '../src/web/engine-client'
import { readLocalImageViaHost } from '../src/web/local-image'

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

describe('Sheets Web Host image adapter', () => {
  it('resolves a token-backed platform file through OfficeHostApi.readFile', async () => {
    const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const pickFile = vi.fn().mockResolvedValue([
      {
        id: 'fs:image-1',
        name: 'platform-image.png',
        mimeType: 'application/octet-stream',
        size: bytes.byteLength,
        version: 'v1',
        transport: 'token',
        token: 'opaque-platform-token',
      },
    ])
    const readFile = vi.fn().mockResolvedValue({
      id: 'fs:image-1',
      name: 'platform-image.png',
      mimeType: 'application/octet-stream',
      size: bytes.byteLength,
      version: 'v1',
      bytes: bytes.buffer as ArrayBuffer,
    })
    const host = { pickFile, readFile } as unknown as OfficeHostApi

    const result = await readLocalImageViaHost(host, { path: 'host-picker://insert-image' })

    expect(result).toEqual({
      mediaType: 'image/png',
      base64: Buffer.from(bytes).toString('base64'),
    })
    expect(pickFile).toHaveBeenCalledWith({
      multiple: false,
      accept: ['image/png', 'image/jpeg', 'image/gif', '.png', '.jpg', '.jpeg', '.gif'],
      mode: 'file',
    })
    expect(readFile).toHaveBeenCalledWith('fs:image-1')
  })

  it('rejects spoofed image metadata after reading the real bytes', async () => {
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04])
    const host = {
      pickFile: vi.fn().mockResolvedValue([
        {
          id: 'fs:not-image',
          name: 'fake.png',
          mimeType: 'image/png',
          size: bytes.byteLength,
          version: 'v1',
          transport: 'token',
          token: 'opaque-platform-token',
        },
      ]),
      readFile: vi.fn().mockResolvedValue({
        id: 'fs:not-image',
        name: 'fake.png',
        mimeType: 'image/png',
        size: bytes.byteLength,
        version: 'v1',
        bytes: bytes.buffer as ArrayBuffer,
      }),
    } as unknown as OfficeHostApi

    await expect(
      readLocalImageViaHost(host, { path: 'host-picker://insert-image' }),
    ).rejects.toThrow('not a PNG/JPEG/GIF image')
  })

  it('rejects oversized platform images before requesting their content', async () => {
    const readFile = vi.fn()
    const host = {
      pickFile: vi.fn().mockResolvedValue([
        {
          id: 'fs:huge-image',
          name: 'huge.png',
          mimeType: 'image/png',
          size: 20 * 1024 * 1024 + 1,
          version: 'v1',
          transport: 'token',
          token: 'opaque-platform-token',
        },
      ]),
      readFile,
    } as unknown as OfficeHostApi

    await expect(
      readLocalImageViaHost(host, { path: 'host-picker://insert-image' }),
    ).rejects.toThrow('exceeds 20MB')
    expect(readFile).not.toHaveBeenCalled()
  })
})
