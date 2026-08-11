import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'

import {
  inventoryXlsx,
  planCellEditsToXlsx,
  type CellEdit,
  type EntrySource,
  type MutationPlan,
} from '../src/gateway/xlsx-gateway'
import type { CellState } from '../src/domain/workbook.types'

const ENGINE_URL = (process.env.XLSX_ENGINE_URL || 'http://127.0.0.1:7301').replace(/\/$/, '')
const MAX_PATCH_ENTRY_BYTES = 256 * 1024 * 1024

interface CorpusCase {
  fixture: string
  sheetName: string
  row: number
  column: number
  after: CellState
}

const CORPUS: readonly CorpusCase[] = [
  {
    fixture: 'compatibility-basic.xlsx',
    sheetName: 'Sheet1',
    row: 0,
    column: 0,
    after: { value: 'Verified by Rust path' },
  },
  {
    fixture: 'compatibility-edit.xlsx',
    sheetName: 'Data',
    row: 0,
    column: 2,
    after: { value: 6 },
  },
  {
    fixture: 'compatibility-structure.xlsx',
    sheetName: 'Data',
    row: 0,
    column: 0,
    after: { value: 99 },
  },
  {
    fixture: 'compatibility-sheets.xlsx',
    sheetName: 'Data',
    row: 0,
    column: 0,
    after: { value: 99 },
  },
  {
    fixture: 'compatibility-kitchen-sink.xlsx',
    sheetName: 'Data',
    row: 0,
    column: 0,
    after: { value: 99 },
  },
]

interface ManifestEntry {
  name: string
  compressedSize: number
  uncompressedSize: number
}

interface OpenedWorkbook {
  sessionId: string
}

interface FixtureReport {
  fixture: string
  passed: boolean
  error?: string
  touchedEntries: string[]
  changedEntries: string[]
  removedEntries: string[]
  addedEntries: string[]
  unexpectedChanges: string[]
  preservedEntryCount: number
}

async function checkedResponse(path: string, init?: RequestInit): Promise<Response> {
  const response = await fetch(`${ENGINE_URL}${path}`, init)
  if (!response.ok) {
    throw new Error(`Engine ${init?.method || 'GET'} ${path} failed (${response.status}): ${await response.text()}`)
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

async function openWorkbook(name: string, bytes: Buffer): Promise<OpenedWorkbook> {
  const response = await checkedResponse(`/v1/workbooks?name=${encodeURIComponent(name)}`, {
    method: 'POST',
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    body: bytes,
  })
  return (await response.json()) as OpenedWorkbook
}

async function manifest(sessionId: string): Promise<ManifestEntry[]> {
  const response = await checkedResponse(
    `/v1/sessions/${encodeURIComponent(sessionId)}/archive/manifest`,
    {
      headers: { Accept: 'application/json', 'X-Xlsx-Session': sessionId },
    },
  )
  const body = (await response.json()) as { entries: ManifestEntry[] }
  return body.entries
}

async function readEntries(
  sessionId: string,
  entries: readonly string[],
): Promise<Map<string, Uint8Array>> {
  const response = await checkedResponse(
    `/v1/sessions/${encodeURIComponent(sessionId)}/archive/read`,
    {
      method: 'POST',
      headers: sessionHeaders(sessionId),
      body: JSON.stringify({ entries }),
    },
  )
  const body = (await response.json()) as {
    entries: Array<{ name: string; contentBase64: string }>
  }
  return new Map(
    body.entries.map((entry) => [entry.name, Uint8Array.from(Buffer.from(entry.contentBase64, 'base64'))]),
  )
}

async function scanEntries(
  sessionId: string,
  entries: readonly string[],
  needle: string,
): Promise<string[]> {
  const response = await checkedResponse(
    `/v1/sessions/${encodeURIComponent(sessionId)}/archive/scan`,
    {
      method: 'POST',
      headers: sessionHeaders(sessionId),
      body: JSON.stringify({ entries, needle }),
    },
  )
  return ((await response.json()) as { matches: string[] }).matches
}

function base64Content(items: ReadonlyMap<string, string | Uint8Array>) {
  return [...items].map(([name, content]) => ({
    name,
    contentBase64:
      typeof content === 'string'
        ? Buffer.from(content, 'utf8').toString('base64')
        : Buffer.from(content).toString('base64'),
  }))
}

function mergeAdditions(plan: MutationPlan): Map<string, string | Uint8Array> {
  const result = new Map<string, string | Uint8Array>()
  for (const [name, content] of plan.added) result.set(name, content)
  for (const [name, content] of plan.addedBinary) result.set(name, content)
  return result
}

async function saveMutation(
  sessionId: string,
  name: string,
  plan: MutationPlan,
): Promise<{ bytes: Buffer; sessionId: string }> {
  const response = await checkedResponse(
    `/v1/sessions/${encodeURIComponent(sessionId)}/archive/save`,
    {
      method: 'POST',
      headers: sessionHeaders(sessionId),
      body: JSON.stringify({
        name,
        replacements: base64Content(plan.replaced),
        removals: plan.removedEntries,
        additions: base64Content(mergeAdditions(plan)),
      }),
    },
  )
  const savedSessionId = response.headers.get('x-xlsx-session')
  if (!savedSessionId) throw new Error('Engine save response did not include x-xlsx-session.')
  return { bytes: Buffer.from(await response.arrayBuffer()), sessionId: savedSessionId }
}

async function deleteSession(sessionId: string): Promise<void> {
  await fetch(`${ENGINE_URL}/v1/sessions/${encodeURIComponent(sessionId)}`, {
    method: 'DELETE',
    headers: { 'X-Xlsx-Session': sessionId },
  })
}

function createEngineEntrySource(
  sessionId: string,
  entries: readonly ManifestEntry[],
): EntrySource {
  const byName = new Map(entries.map((entry) => [entry.name, entry]))
  const textCache = new Map<string, string>()
  const decoder = new TextDecoder()

  return {
    paths: async () => entries.map((entry) => entry.name),
    has: async (path) => byName.has(path),
    canPatch: async (path) => (byName.get(path)?.uncompressedSize ?? 0) <= MAX_PATCH_ENTRY_BYTES,
    containsText: async (path, needle) => (await scanEntries(sessionId, [path], needle)).includes(path),
    readText: async (path) => {
      const cached = textCache.get(path)
      if (cached !== undefined) return cached
      const entry = byName.get(path)
      if (!entry) throw new Error(`Workbook is missing ${path}.`)
      if (entry.uncompressedSize > MAX_PATCH_ENTRY_BYTES) {
        throw new Error(`${path} is too large to patch in the compatibility runner.`)
      }
      const content = (await readEntries(sessionId, [path])).get(path)
      if (!content) throw new Error(`Engine did not return ${path}.`)
      const text = decoder.decode(content)
      textCache.set(path, text)
      return text
    },
  }
}

async function planSingleCellEdit(source: EntrySource, entry: CorpusCase): Promise<MutationPlan> {
  const edit: CellEdit = {
    sheetName: entry.sheetName,
    row: entry.row,
    column: entry.column,
    writeValue: true,
    cell: entry.after,
  }

  return planCellEditsToXlsx(
    source,
    [edit],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    [],
  )
}

async function verifyCase(entry: CorpusCase): Promise<FixtureReport> {
  const sourceBytes = await readFile(resolve('fixtures/generated', entry.fixture))
  let sourceSessionId: string | null = null
  let savedSessionId: string | null = null

  try {
    const opened = await openWorkbook(entry.fixture, sourceBytes)
    sourceSessionId = opened.sessionId
    const entries = await manifest(sourceSessionId)
    const source = createEngineEntrySource(sourceSessionId, entries)
    const plan = await planSingleCellEdit(source, entry)
    const saved = await saveMutation(sourceSessionId, entry.fixture, plan)
    savedSessionId = saved.sessionId

    const beforeEntries = await inventoryXlsx(sourceBytes)
    const afterEntries = await inventoryXlsx(saved.bytes)
    const beforeByPath = new Map(beforeEntries.map((item) => [item.path, item.sha256]))
    const afterPaths = new Set(afterEntries.map((item) => item.path))
    const changedEntries = afterEntries
      .filter((item) => beforeByPath.get(item.path) !== item.sha256)
      .map((item) => item.path)
    const removedEntries = beforeEntries
      .filter((item) => !afterPaths.has(item.path))
      .map((item) => item.path)
    const beforePaths = new Set(beforeEntries.map((item) => item.path))
    const addedEntries = afterEntries.filter((item) => !beforePaths.has(item.path)).map((item) => item.path)

    const allowed = new Set([
      ...plan.touchedEntries,
      ...plan.removedEntries,
      ...plan.addedEntries,
    ])
    const unexpectedChanges = [...changedEntries, ...removedEntries, ...addedEntries].filter(
      (path) => !allowed.has(path),
    )

    return {
      fixture: entry.fixture,
      passed: unexpectedChanges.length === 0 && changedEntries.length > 0,
      touchedEntries: [...plan.touchedEntries],
      changedEntries,
      removedEntries,
      addedEntries,
      unexpectedChanges,
      preservedEntryCount: afterEntries.filter((item) => beforeByPath.get(item.path) === item.sha256).length,
    }
  } catch (error) {
    return {
      fixture: entry.fixture,
      passed: false,
      error: error instanceof Error ? error.message : String(error),
      touchedEntries: [],
      changedEntries: [],
      removedEntries: [],
      addedEntries: [],
      unexpectedChanges: [],
      preservedEntryCount: 0,
    }
  } finally {
    if (savedSessionId) await deleteSession(savedSessionId)
    if (sourceSessionId) await deleteSession(sourceSessionId)
  }
}

async function main(): Promise<void> {
  await checkedResponse('/health')

  const fixtures: FixtureReport[] = []
  for (const entry of CORPUS) fixtures.push(await verifyCase(entry))
  const report = {
    engineUrl: ENGINE_URL,
    passed: fixtures.every((fixture) => fixture.passed),
    fixtureCount: fixtures.length,
    fixtures,
  }

  await mkdir(resolve('reports'), { recursive: true })
  await writeFile(
    resolve('reports/web-engine-compatibility.json'),
    `${JSON.stringify(report, null, 2)}\n`,
  )
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`)
  if (!report.passed) process.exitCode = 1
}

void main()
