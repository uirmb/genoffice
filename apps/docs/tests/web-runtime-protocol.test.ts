import { describe, expect, it, vi } from 'vitest'
import type { EditorIframeBridge } from '@genoffice/web-runtime'
import { EmbeddedOfficeHost } from '@genoffice/web-runtime'
import type { EditorToHostMessage, HostToEditorMessage } from '@genoffice/office-protocol'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

function bytes(...values: number[]): ArrayBuffer {
  return new Uint8Array(values).buffer
}

function createBridge() {
  const requests: EditorToHostMessage[] = []
  const sends: EditorToHostMessage[] = []

  const request = vi.fn(async (message: EditorToHostMessage, expectedType: string) => {
    requests.push(message)

    if (message.type === 'office:pick-assets') {
      expect(expectedType).toBe('office:pick-assets-result')
      return {
        protocol: 1,
        type: 'office:pick-assets-result',
        requestId: message.requestId,
        payload: {
          status: 'selected',
          files: [
            {
              id: 'asset-1',
              name: 'image.png',
              mimeType: 'image/png',
              size: 3,
              version: 1,
              bytes: bytes(1, 2, 3),
              transport: 'buffer',
            },
          ],
        },
      } satisfies HostToEditorMessage
    }

    if (message.type === 'office:pick-document') {
      expect(expectedType).toBe('office:pick-document-result')
      return {
        protocol: 1,
        type: 'office:pick-document-result',
        requestId: message.requestId,
        payload: {
          status: 'selected',
          selectionId: 'selection-1',
          file: {
            id: 'doc-1',
            nodeId: 'doc-1',
            name: 'picked.docx',
            mimeType: DOCX_MIME,
            size: 3,
            version: 7,
            bytes: bytes(4, 5, 6),
            transport: 'buffer',
          },
        },
      } satisfies HostToEditorMessage
    }

    if (message.type === 'office:document-opened') {
      expect(expectedType).toBe('office:document-opened-result')
      return {
        protocol: 1,
        type: 'office:document-opened-result',
        requestId: message.requestId,
        payload: {
          ok: true,
          file: {
            id: 'doc-1',
            nodeId: 'doc-1',
            name: 'picked.docx',
            mimeType: DOCX_MIME,
            size: 3,
            version: 7,
            transport: 'buffer',
          },
        },
      } satisfies HostToEditorMessage
    }

    if (message.type === 'office:save-document') {
      expect(expectedType).toBe('office:save-document-result')
      return {
        protocol: 1,
        type: 'office:save-document-result',
        requestId: message.requestId,
        payload: {
          ok: true,
          file: {
            id: message.payload.file.id,
            name: message.payload.file.name,
            mimeType: message.payload.file.mimeType,
            size: message.payload.bytes.byteLength,
            version: 8,
            transport: 'buffer',
          },
        },
      } satisfies HostToEditorMessage
    }

    if (message.type === 'office:download-document') {
      expect(expectedType).toBe('office:download-document-result')
      return {
        protocol: 1,
        type: 'office:download-document-result',
        requestId: message.requestId,
        payload: { ok: true },
      } satisfies HostToEditorMessage
    }

    throw new Error(`Unexpected request: ${message.type}`)
  })

  const send = vi.fn((message: EditorToHostMessage) => {
    sends.push(message)
  })

  return {
    bridge: { request, send } as unknown as EditorIframeBridge,
    requests,
    sends,
  }
}

describe('EmbeddedOfficeHost stable wire protocol compatibility', () => {
  it('translates legacy image pickFile calls to office:pick-assets with buffer results', async () => {
    const { bridge, requests } = createBridge()
    const host = new EmbeddedOfficeHost(bridge, 'en-US')

    const selected = await host.pickFile({
      multiple: false,
      accept: ['image/png', '.png'],
      mode: 'file',
    })

    expect(selected?.[0]?.transport).toBe('buffer')
    expect(selected?.[0]?.bytes).toBeInstanceOf(ArrayBuffer)
    expect(requests.map((message) => message.type)).toEqual(['office:pick-assets'])
  })

  it('translates legacy document pickFile into pick-document and binds before save', async () => {
    const { bridge, requests } = createBridge()
    const host = new EmbeddedOfficeHost(bridge, 'en-US')

    const selected = await host.pickFile({
      multiple: false,
      accept: [DOCX_MIME, '.docx'],
      mode: 'file',
    })
    expect(selected?.[0]?.id).toBe('doc-1')
    expect(requests.map((message) => message.type)).toEqual(['office:pick-document'])

    // Existing editors call setTitle after they have accepted the selected file.
    // During protocol-v1 migration the runtime uses that point to submit the
    // pending selection. saveDocument must wait for this bind transaction.
    host.setTitle('picked.docx')

    const result = await host.saveDocument({
      file: {
        id: 'doc-1',
        name: 'picked.docx',
        mimeType: DOCX_MIME,
        size: 3,
        version: 7,
        transport: 'buffer',
      },
      bytes: bytes(9, 9, 9),
      baseVersion: 7,
      mode: 'save',
    })

    expect(result.ok).toBe(true)
    expect(requests.map((message) => message.type)).toEqual([
      'office:pick-document',
      'office:document-opened',
      'office:save-document',
    ])
    expect(requests.some((message) => message.type === 'office:pick-file')).toBe(false)
    expect(requests.some((message) => message.type === 'office:read-file')).toBe(false)
  })

  it('translates legacy exportDocument to office:download-document', async () => {
    const { bridge, requests } = createBridge()
    const host = new EmbeddedOfficeHost(bridge, 'en-US')

    const result = await host.exportDocument({
      format: 'docx',
      file: {
        id: 'doc-1',
        name: 'download.docx',
        mimeType: DOCX_MIME,
        size: 2,
        version: null,
        transport: 'buffer',
      },
      bytes: bytes(1, 2),
    })

    expect(result.ok).toBe(true)
    expect(requests.map((message) => message.type)).toEqual(['office:download-document'])
    expect(requests.some((message) => message.type === 'office:export-document')).toBe(false)
  })

  it('translates legacy requestClose to office:close-approved', async () => {
    const { bridge, sends } = createBridge()
    const host = new EmbeddedOfficeHost(bridge, 'en-US')

    await host.requestClose()

    expect(sends).toHaveLength(1)
    expect(sends[0]?.type).toBe('office:close-approved')
    if (sends[0]?.type === 'office:close-approved') {
      expect(sends[0].payload.reason).toBe('file-menu')
      expect(sends[0].requestId).toBeTruthy()
    }
  })
})
