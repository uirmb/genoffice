import {
  workbookFileSchema,
  workbookFormulaCellsResultSchema,
  workbookMediaResultSchema,
  workbookRangeResultSchema,
  workbookRecalcResultSchema,
  type WorkbookFile,
  type WorkbookFormulaCellsRequest,
  type WorkbookFormulaCellsResult,
  type WorkbookMediaRequest,
  type WorkbookMediaResult,
  type WorkbookRangeRequest,
  type WorkbookRangeResult,
  type WorkbookRecalcRequest,
  type WorkbookRecalcResult,
} from '../shared/desktop-api'

const ENGINE_BASE = '/xlsx-engine'
const MAX_MEDIA_BYTES = 20 * 1024 * 1024

export interface XlsxEngineHealth {
  ok: boolean
  service: string
  sessionStore: string
}

export interface XlsxEngineSession {
  sessionId: string
  source: 'blank' | 'uploaded'
}

export interface XlsxArchiveEntry {
  name: string
  crc32: number
  compressedSize: number
  uncompressedSize: number
}

export interface XlsxArchiveMutation {
  replacements: ReadonlyMap<string, string | Uint8Array>
  removals: readonly string[]
  additions: ReadonlyMap<string, string | Uint8Array>
}

export interface SavedXlsxWorkbook {
  file: WorkbookFile
  bytes: ArrayBuffer
}

async function readJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const detail = await response.text().catch(() => '')
    throw new Error(
      `XLSX engine request failed (${response.status})${detail ? `: ${detail}` : ''}`,
    )
  }
  return (await response.json()) as T
}

async function assertOk(response: Response): Promise<Response> {
  if (!response.ok) {
    const detail = await response.text().catch(() => '')
    throw new Error(
      `XLSX engine request failed (${response.status})${detail ? `: ${detail}` : ''}`,
    )
  }
  return response
}

function sessionHeaders(sessionId: string): HeadersInit {
  return {
    Accept: 'application/json',
    'Content-Type': 'application/json',
    'X-Xlsx-Session': sessionId,
  }
}

function bytesToBase64(value: string | Uint8Array): string {
  const bytes = typeof value === 'string' ? new TextEncoder().encode(value) : value
  let binary = ''
  const chunkSize = 0x8000
  for (let offset = 0; offset < bytes.byteLength; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize))
  }
  return btoa(binary)
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

function archiveContent(items: ReadonlyMap<string, string | Uint8Array>) {
  return [...items].map(([name, content]) => ({
    name,
    contentBase64: bytesToBase64(content),
  }))
}

function workbookSheetNames(workbook: WorkbookFile): {
  namesById: Map<string, string>
  idsByName: Map<string, string>
} {
  return {
    namesById: new Map(workbook.sheets.map((sheet) => [sheet.id, sheet.name])),
    idsByName: new Map(workbook.sheets.map((sheet) => [sheet.name, sheet.id])),
  }
}

function mediaTypeForPath(path: string): string | null {
  const extension = path.split('.').at(-1)?.toLowerCase()
  if (extension === 'png') return 'image/png'
  if (extension === 'jpg' || extension === 'jpeg') return 'image/jpeg'
  if (extension === 'gif') return 'image/gif'
  if (extension === 'webp') return 'image/webp'
  if (extension === 'bmp') return 'image/bmp'
  if (extension === 'svg') return 'image/svg+xml'
  return null
}

export async function getXlsxEngineHealth(): Promise<XlsxEngineHealth> {
  const response = await fetch(`${ENGINE_BASE}/health`, {
    headers: { Accept: 'application/json' },
  })
  return readJson<XlsxEngineHealth>(response)
}

export async function createBlankXlsxSession(): Promise<XlsxEngineSession> {
  const response = await fetch(`${ENGINE_BASE}/v1/sessions`, {
    method: 'POST',
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ source: 'blank' }),
  })
  return readJson<XlsxEngineSession>(response)
}

export async function createBlankXlsxWorkbook(name = 'Untitled.xlsx'): Promise<WorkbookFile> {
  const response = await fetch(
    `${ENGINE_BASE}/v1/workbooks/blank?name=${encodeURIComponent(name)}`,
    {
      method: 'POST',
      headers: { Accept: 'application/json' },
    },
  )
  return workbookFileSchema.parse(await readJson<unknown>(response))
}

export async function openXlsxWorkbookBytes(name: string, bytes: ArrayBuffer): Promise<WorkbookFile> {
  const response = await fetch(
    `${ENGINE_BASE}/v1/workbooks?name=${encodeURIComponent(name || 'workbook.xlsx')}`,
    {
      method: 'POST',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: bytes,
    },
  )
  return workbookFileSchema.parse(await readJson<unknown>(response))
}

export async function openXlsxWorkbook(file: File): Promise<WorkbookFile> {
  return openXlsxWorkbookBytes(file.name, await file.arrayBuffer())
}

export async function getXlsxWorkbookMetadata(sessionId: string): Promise<WorkbookFile> {
  const response = await fetch(`${ENGINE_BASE}/v1/sessions/${encodeURIComponent(sessionId)}`, {
    headers: {
      Accept: 'application/json',
      'X-Xlsx-Session': sessionId,
    },
  })
  return workbookFileSchema.parse(await readJson<unknown>(response))
}

export async function readXlsxWorkbookRange(
  request: WorkbookRangeRequest,
): Promise<WorkbookRangeResult> {
  const response = await fetch(
    `${ENGINE_BASE}/v1/sessions/${encodeURIComponent(request.sessionId)}/ranges`,
    {
      method: 'POST',
      headers: sessionHeaders(request.sessionId),
      body: JSON.stringify({
        sheetId: request.sheetId,
        range: request.range,
      }),
    },
  )
  return workbookRangeResultSchema.parse(await readJson<unknown>(response))
}

export async function readXlsxWorkbookFormulaCells(
  request: WorkbookFormulaCellsRequest,
): Promise<WorkbookFormulaCellsResult> {
  const response = await fetch(
    `${ENGINE_BASE}/v1/sessions/${encodeURIComponent(request.sessionId)}/formulas`,
    {
      method: 'POST',
      headers: sessionHeaders(request.sessionId),
      body: JSON.stringify({ sheetId: request.sheetId }),
    },
  )
  return workbookFormulaCellsResultSchema.parse(await readJson<unknown>(response))
}

export async function recalcXlsxWorkbook(
  request: WorkbookRecalcRequest,
  workbook: WorkbookFile,
): Promise<WorkbookRecalcResult> {
  const { namesById, idsByName } = workbookSheetNames(workbook)
  const sheetName = (sheetId: string): string => {
    const name = namesById.get(sheetId)
    if (!name) throw new Error(`Unknown worksheet ${sheetId}.`)
    return name
  }

  const response = await fetch(
    `${ENGINE_BASE}/v1/sessions/${encodeURIComponent(request.sessionId)}/recalc`,
    {
      method: 'POST',
      headers: sessionHeaders(request.sessionId),
      body: JSON.stringify({
        edits: request.edits.map((edit) => ({
          sheet: sheetName(edit.sheetId),
          row: edit.row,
          column: edit.column,
          input: edit.input,
        })),
        reads: request.reads.map((read) => ({
          sheet: sheetName(read.sheetId),
          range: read.range,
        })),
      }),
    },
  )
  const result = await readJson<{
    cells: Array<{
      sheet: string
      row: number
      column: number
      formatted: string
      number?: number
      isFormula: boolean
    }>
  }>(response)

  return workbookRecalcResultSchema.parse({
    cells: result.cells.map(({ sheet, ...cell }) => {
      const sheetId = idsByName.get(sheet)
      if (!sheetId) throw new Error(`XLSX engine returned unknown worksheet ${sheet}.`)
      return { sheetId, ...cell }
    }),
  })
}

export async function readXlsxWorkbookMedia(
  request: WorkbookMediaRequest,
  workbook: WorkbookFile,
): Promise<WorkbookMediaResult> {
  const visual = workbook.visuals.find((candidate) => candidate.id === request.visualId)
  if (!visual?.mediaPath) throw new Error(`Unknown workbook image ${request.visualId}.`)

  const manifest = await getXlsxArchiveManifest(request.sessionId)
  const entry = manifest.find((candidate) => candidate.name === visual.mediaPath)
  if (!entry) throw new Error(`Workbook is missing ${visual.mediaPath}.`)
  if (entry.uncompressedSize > MAX_MEDIA_BYTES) {
    throw new Error(`Workbook image exceeds the ${MAX_MEDIA_BYTES / 1024 / 1024}MB preview limit.`)
  }

  const entries = await readXlsxArchiveEntries(request.sessionId, [visual.mediaPath])
  const bytes = entries.get(visual.mediaPath)
  if (!bytes) throw new Error(`XLSX Engine did not return ${visual.mediaPath}.`)

  const mediaType = visual.mediaType ?? mediaTypeForPath(visual.mediaPath)
  if (!mediaType?.startsWith('image/')) {
    throw new Error(`Unsupported workbook image type for ${visual.mediaPath}.`)
  }

  return workbookMediaResultSchema.parse({
    mediaType,
    base64: bytesToBase64(bytes),
  })
}

export async function getXlsxArchiveManifest(sessionId: string): Promise<XlsxArchiveEntry[]> {
  const response = await fetch(
    `${ENGINE_BASE}/v1/sessions/${encodeURIComponent(sessionId)}/archive/manifest`,
    {
      headers: {
        Accept: 'application/json',
        'X-Xlsx-Session': sessionId,
      },
    },
  )
  const body = await readJson<{ entries: XlsxArchiveEntry[] }>(response)
  return body.entries
}

export async function readXlsxArchiveEntries(
  sessionId: string,
  entries: readonly string[],
): Promise<Map<string, Uint8Array>> {
  const response = await fetch(
    `${ENGINE_BASE}/v1/sessions/${encodeURIComponent(sessionId)}/archive/read`,
    {
      method: 'POST',
      headers: sessionHeaders(sessionId),
      body: JSON.stringify({ entries }),
    },
  )
  const body = await readJson<{
    entries: { name: string; contentBase64: string }[]
  }>(response)
  return new Map(body.entries.map((entry) => [entry.name, base64ToBytes(entry.contentBase64)]))
}

export async function scanXlsxArchiveEntries(
  sessionId: string,
  entries: readonly string[],
  needle: string,
): Promise<string[]> {
  const response = await fetch(
    `${ENGINE_BASE}/v1/sessions/${encodeURIComponent(sessionId)}/archive/scan`,
    {
      method: 'POST',
      headers: sessionHeaders(sessionId),
      body: JSON.stringify({ entries, needle }),
    },
  )
  const body = await readJson<{ matches: string[] }>(response)
  return body.matches
}

export async function saveXlsxArchiveMutation(
  sessionId: string,
  name: string,
  mutation: XlsxArchiveMutation,
): Promise<SavedXlsxWorkbook> {
  const response = await assertOk(
    await fetch(
      `${ENGINE_BASE}/v1/sessions/${encodeURIComponent(sessionId)}/archive/save`,
      {
        method: 'POST',
        headers: sessionHeaders(sessionId),
        body: JSON.stringify({
          name,
          replacements: archiveContent(mutation.replacements),
          removals: mutation.removals,
          additions: archiveContent(mutation.additions),
        }),
      },
    ),
  )
  const savedSessionId = response.headers.get('x-xlsx-session')
  if (!savedSessionId) throw new Error('XLSX engine save response did not include a session id.')
  const bytes = await response.arrayBuffer()
  const file = await getXlsxWorkbookMetadata(savedSessionId)
  return { file, bytes }
}

export async function deleteXlsxSession(sessionId: string): Promise<void> {
  const response = await fetch(`${ENGINE_BASE}/v1/sessions/${encodeURIComponent(sessionId)}`, {
    method: 'DELETE',
    headers: {
      'X-Xlsx-Session': sessionId,
    },
  })
  if (response.status !== 204 && response.status !== 404) {
    throw new Error(`Unable to release XLSX session (${response.status}).`)
  }
}
