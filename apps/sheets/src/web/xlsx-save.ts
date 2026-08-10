import type { WorkbookFile, WorkbookSaveRequest } from '../shared/desktop-api'
import {
  planCellEditsToXlsx,
  type CellEdit,
  type EntrySource,
  type MutationPlan,
} from '../gateway/xlsx-gateway'
import {
  getXlsxArchiveManifest,
  readXlsxArchiveEntries,
  saveXlsxArchiveMutation,
  scanXlsxArchiveEntries,
  type SavedXlsxWorkbook,
  type XlsxArchiveEntry,
} from './engine-client'

const MAX_PATCH_ENTRY_BYTES = 256 * 1024 * 1024

function createEngineEntrySource(
  sessionId: string,
  manifest: readonly XlsxArchiveEntry[],
): EntrySource {
  const entryByName = new Map(manifest.map((entry) => [entry.name, entry]))
  const textCache = new Map<string, string>()
  const decoder = new TextDecoder()

  return {
    paths: async () => manifest.map((entry) => entry.name),
    has: async (path) => entryByName.has(path),
    canPatch: async (path) =>
      (entryByName.get(path)?.uncompressedSize ?? 0) <= MAX_PATCH_ENTRY_BYTES,
    containsText: async (path, needle) => {
      const matches = await scanXlsxArchiveEntries(sessionId, [path], needle)
      return matches.includes(path)
    },
    readText: async (path) => {
      const cached = textCache.get(path)
      if (cached !== undefined) return cached
      const entry = entryByName.get(path)
      if (!entry) throw new Error(`Workbook is missing ${path}.`)
      if (entry.uncompressedSize > MAX_PATCH_ENTRY_BYTES) {
        throw new Error(
          `${path} is ${entry.uncompressedSize} bytes uncompressed — too large to edit. ` +
            'Entries above 256MB can be preserved but not patched.',
        )
      }
      const entries = await readXlsxArchiveEntries(sessionId, [path])
      const bytes = entries.get(path)
      if (!bytes) throw new Error(`XLSX Engine did not return ${path}.`)
      const content = decoder.decode(bytes)
      textCache.set(path, content)
      return content
    },
  }
}

function mergeAdditions(plan: MutationPlan): Map<string, string | Uint8Array> {
  const additions = new Map<string, string | Uint8Array>()
  for (const [path, content] of plan.added) additions.set(path, content)
  for (const [path, content] of plan.addedBinary) additions.set(path, content)
  return additions
}

function hasUnsupportedAdvancedEdits(request: WorkbookSaveRequest): boolean {
  return (
    request.structuralOps.length > 0 ||
    request.chartEdits.length > 0 ||
    request.visualEdits.length > 0 ||
    request.visualAdditions.length > 0 ||
    request.tableAdditions.length > 0 ||
    request.pivotAdditions.length > 0 ||
    request.sheetOps.length > 0 ||
    request.filterStates.length > 0 ||
    request.hyperlinkEdits.length > 0 ||
    request.cfStates.length > 0 ||
    request.dvStates.length > 0 ||
    request.pageSetupStates.length > 0 ||
    request.noteStates.length > 0 ||
    request.pivotCacheRefreshPaths.length > 0 ||
    request.pivotRefreshUpdates.length > 0 ||
    request.sheetProtections.length > 0 ||
    request.sparklineAdditions.length > 0 ||
    request.formulaValues.length > 0 ||
    request.definedNamesState !== null
  )
}

function toPlannerCellEdits(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): CellEdit[] {
  const names = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet.name]))
  const sheetName = (sheetId: string): string => {
    const name = names.get(sheetId)
    if (!name) throw new Error(`Unknown worksheet ${sheetId}.`)
    return name
  }

  return request.edits.map((edit) => ({
    sheetName: sheetName(edit.sheetId),
    row: edit.row,
    column: edit.column,
    writeValue: edit.writeValue,
    cell: {
      value: edit.value,
      ...(edit.formula === undefined ? {} : { formula: edit.formula }),
    },
    ...(edit.style === undefined ? {} : { style: edit.style }),
    ...(edit.rich === undefined ? {} : { rich: edit.rich }),
    ...(edit.styleReset === undefined ? {} : { styleReset: edit.styleReset }),
  }))
}

/**
 * Browser preservation-save path. The same mutation planner used by Electron
 * reads package parts through the Rust session API; Rust only reassembles the
 * planned replacement/add/remove sets and returns standard XLSX bytes.
 *
 * The first Web milestone intentionally enables the common cell-edit path
 * before the higher-level journal families. Unsupported advanced journals fail
 * closed rather than producing a partially saved workbook.
 */
export async function saveWorkbookRequestViaEngine(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
  name: string,
): Promise<SavedXlsxWorkbook & { touchedEntries: readonly string[] }> {
  if (hasUnsupportedAdvancedEdits(request)) {
    throw new Error('This workbook contains edits that are not enabled in Sheets Web save yet.')
  }

  const manifest = await getXlsxArchiveManifest(request.sessionId)
  const source = createEngineEntrySource(request.sessionId, manifest)
  const plan = await planCellEditsToXlsx(source, toPlannerCellEdits(request, workbook))

  const saved = await saveXlsxArchiveMutation(request.sessionId, name, {
    replacements: plan.replaced,
    removals: plan.removedEntries,
    additions: mergeAdditions(plan),
  })

  return {
    ...saved,
    touchedEntries: plan.touchedEntries,
  }
}
