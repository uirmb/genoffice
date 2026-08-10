import {
  workbookFileSchema,
  workbookRangeResultSchema,
  type WorkbookFile,
  type WorkbookRangeRequest,
  type WorkbookRangeResult,
} from '../shared/desktop-api'

const ENGINE_BASE = '/xlsx-engine'

export interface XlsxEngineHealth {
  ok: boolean
  service: string
  sessionStore: string
}

export interface XlsxEngineSession {
  sessionId: string
  source: 'blank' | 'uploaded'
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

function sessionHeaders(sessionId: string): HeadersInit {
  return {
    Accept: 'application/json',
    'Content-Type': 'application/json',
    'X-Xlsx-Session': sessionId,
  }
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
