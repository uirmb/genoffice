import type { WorkbookFile, WorkbookSaveRequest } from '../shared/desktop-api'
import {
  planCellEditsToXlsx,
  type CellEdit,
  type EntrySource,
  type MutationPlan,
  type SheetFormulaValues,
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

type PlannerFilterStates = Parameters<typeof planCellEditsToXlsx>[5]
type PlannerHyperlinkEdits = Parameters<typeof planCellEditsToXlsx>[6]
type PlannerCfStates = Parameters<typeof planCellEditsToXlsx>[7]
type PlannerDvStates = Parameters<typeof planCellEditsToXlsx>[8]
type PlannerSheetProtections = Parameters<typeof planCellEditsToXlsx>[9]
type PlannerPageSetupStates = Parameters<typeof planCellEditsToXlsx>[12]
type PlannerNoteStates = Parameters<typeof planCellEditsToXlsx>[13]

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
    request.pivotCacheRefreshPaths.length > 0 ||
    request.pivotRefreshUpdates.length > 0 ||
    request.sparklineAdditions.length > 0 ||
    request.definedNamesState !== null
  )
}

function workbookSheetNames(workbook: WorkbookFile): Map<string, string> {
  return new Map(workbook.sheets.map((sheet) => [sheet.id, sheet.name]))
}

function requiredSheetName(names: ReadonlyMap<string, string>, sheetId: string): string {
  const name = names.get(sheetId)
  if (!name) throw new Error(`Unknown worksheet ${sheetId}.`)
  return name
}

function toPlannerCellEdits(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): CellEdit[] {
  const names = workbookSheetNames(workbook)

  return request.edits.map((edit) => ({
    sheetName: requiredSheetName(names, edit.sheetId),
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

function toPlannerFilterStates(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): PlannerFilterStates {
  const names = workbookSheetNames(workbook)
  return request.filterStates.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    filter: state.filter,
    hiddenRows: state.hiddenRows,
    visibilityRange: state.visibilityRange,
  }))
}

function toPlannerHyperlinkEdits(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): PlannerHyperlinkEdits {
  const names = workbookSheetNames(workbook)
  const linksBySheet = new Map<
    string,
    Array<{ row: number; column: number; target: string | null }>
  >()

  for (const link of request.hyperlinkEdits) {
    const sheetName = requiredSheetName(names, link.sheetId)
    const links = linksBySheet.get(sheetName) ?? []
    links.push({ row: link.row, column: link.column, target: link.target })
    linksBySheet.set(sheetName, links)
  }

  return [...linksBySheet].map(([sheetName, links]) => ({ sheetName, edits: links }))
}

function toPlannerCfStates(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): PlannerCfStates {
  const names = workbookSheetNames(workbook)
  return request.cfStates.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    rules: state.rules,
  }))
}

function toPlannerDvStates(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): PlannerDvStates {
  const names = workbookSheetNames(workbook)
  return request.dvStates.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    rules: state.rules,
  }))
}

function toPlannerSheetProtections(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): PlannerSheetProtections {
  const names = workbookSheetNames(workbook)
  return request.sheetProtections.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    protected: state.protected,
  }))
}

function toPlannerPageSetupStates(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): PlannerPageSetupStates {
  const names = workbookSheetNames(workbook)
  return request.pageSetupStates.map(({ sheetId, ...state }) => ({
    sheetName: requiredSheetName(names, sheetId),
    ...state,
  }))
}

function toPlannerNoteStates(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): PlannerNoteStates {
  const names = workbookSheetNames(workbook)
  return request.noteStates.map(({ sheetId, notes }) => ({
    sheetName: requiredSheetName(names, sheetId),
    notes,
  }))
}

function toPlannerFormulaValues(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): SheetFormulaValues[] {
  const names = workbookSheetNames(workbook)
  const valuesBySheet = new Map<string, SheetFormulaValues['cells'][number][]>()

  for (const value of request.formulaValues) {
    const sheetName = requiredSheetName(names, value.sheetId)
    const cells = valuesBySheet.get(sheetName) ?? []
    cells.push({
      row: value.row,
      column: value.column,
      value: value.value,
    })
    valuesBySheet.set(sheetName, cells)
  }

  return [...valuesBySheet].map(([sheetName, cells]) => ({ sheetName, cells }))
}

/**
 * Browser preservation-save path. The same mutation planner used by Electron
 * reads package parts through the Rust session API; Rust only reassembles the
 * planned replacement/add/remove sets and returns standard XLSX bytes.
 *
 * Web enables journal families only after they have direct preservation tests.
 * Unsupported high-complexity journals still fail closed rather than producing
 * a partially saved workbook.
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
  const plan = await planCellEditsToXlsx(
    source,
    toPlannerCellEdits(request, workbook),
    [],
    [],
    undefined,
    toPlannerFilterStates(request, workbook),
    toPlannerHyperlinkEdits(request, workbook),
    toPlannerCfStates(request, workbook),
    toPlannerDvStates(request, workbook),
    toPlannerSheetProtections(request, workbook),
    null,
    [],
    toPlannerPageSetupStates(request, workbook),
    toPlannerNoteStates(request, workbook),
    [],
    [],
    [],
    [],
    [],
    [],
    toPlannerFormulaValues(request, workbook),
  )

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
