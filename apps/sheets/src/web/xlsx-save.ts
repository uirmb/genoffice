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

type PlannerStructuralOps = NonNullable<Parameters<typeof planCellEditsToXlsx>[2]>
type PlannerStructuralOp = PlannerStructuralOps[number]['ops'][number]
type PlannerSheetPlan = NonNullable<Parameters<typeof planCellEditsToXlsx>[4]>
type PlannerFilterStates = Parameters<typeof planCellEditsToXlsx>[5]
type PlannerHyperlinkEdits = Parameters<typeof planCellEditsToXlsx>[6]
type PlannerCfStates = Parameters<typeof planCellEditsToXlsx>[7]
type PlannerDvStates = Parameters<typeof planCellEditsToXlsx>[8]
type PlannerSheetProtections = Parameters<typeof planCellEditsToXlsx>[9]
type PlannerPageSetupStates = Parameters<typeof planCellEditsToXlsx>[12]
type PlannerNoteStates = Parameters<typeof planCellEditsToXlsx>[13]

interface SheetPlanContext {
  readonly plan: PlannerSheetPlan | undefined
  /** Planner-stage name: original file name for existing sheets, final name for additions. */
  readonly plannerNames: ReadonlyMap<string, string>
}

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
    request.visualEdits.length > 0 ||
    request.visualAdditions.length > 0 ||
    request.tableAdditions.length > 0 ||
    request.pivotAdditions.length > 0 ||
    request.pivotCacheRefreshPaths.length > 0 ||
    request.pivotRefreshUpdates.length > 0 ||
    request.sparklineAdditions.length > 0
  )
}

function originalSheetNames(workbook: WorkbookFile): Map<string, string> {
  return new Map(workbook.sheets.map((sheet) => [sheet.id, sheet.name]))
}

function requiredSheetName(names: ReadonlyMap<string, string>, sheetId: string): string {
  const name = names.get(sheetId)
  if (!name) throw new Error(`Unknown worksheet ${sheetId}.`)
  return name
}

/**
 * Collapses the renderer's ordered sheet journal into the gateway's declarative
 * SheetEditPlan. Existing sheets keep their original file name until the final
 * sheet-surgery phase; newly added sheets use their final name from the start
 * because their worksheet part is allocated under that name.
 */
function buildSheetPlanContext(
  request: WorkbookSaveRequest,
  workbook: WorkbookFile,
): SheetPlanContext {
  const originalNames = originalSheetNames(workbook)
  if (request.sheetOps.length === 0) {
    return { plan: undefined, plannerNames: originalNames }
  }

  const finalNames = new Map(originalNames)
  const added = new Map<
    string,
    { name: string; sourceSheetId?: string | undefined; sequence: number }
  >()
  const removed = new Set<string>()
  const hidden = new Map<string, boolean>()
  let additionSequence = 0
  let orderChanged = false

  for (const op of request.sheetOps) {
    switch (op.kind) {
      case 'rename-sheet': {
        if (!finalNames.has(op.sheetId)) throw new Error(`Unknown worksheet ${op.sheetId}.`)
        finalNames.set(op.sheetId, op.newName)
        const addition = added.get(op.sheetId)
        if (addition) addition.name = op.newName
        break
      }
      case 'add-sheet': {
        if (finalNames.has(op.sheetId)) {
          throw new Error(`Worksheet id ${op.sheetId} already exists.`)
        }
        finalNames.set(op.sheetId, op.name)
        added.set(op.sheetId, { name: op.name, sequence: additionSequence++ })
        break
      }
      case 'duplicate-sheet': {
        if (finalNames.has(op.sheetId)) {
          throw new Error(`Worksheet id ${op.sheetId} already exists.`)
        }
        if (!originalNames.has(op.sourceSheetId)) {
          throw new Error(
            'Duplicating a sheet that was itself added in the current unsaved session is not supported yet.',
          )
        }
        finalNames.set(op.sheetId, op.name)
        added.set(op.sheetId, {
          name: op.name,
          sourceSheetId: op.sourceSheetId,
          sequence: additionSequence++,
        })
        break
      }
      case 'remove-sheet': {
        if (!finalNames.has(op.sheetId)) throw new Error(`Unknown worksheet ${op.sheetId}.`)
        removed.add(op.sheetId)
        break
      }
      case 'set-sheet-hidden': {
        if (!finalNames.has(op.sheetId)) throw new Error(`Unknown worksheet ${op.sheetId}.`)
        hidden.set(op.sheetId, op.hidden)
        break
      }
      case 'reorder-sheets':
        orderChanged = true
        break
    }
  }

  // An added-then-removed sheet never needs to enter the package.
  const canceledAdditionIds = [...removed].filter((sheetId) => added.has(sheetId))
  for (const sheetId of canceledAdditionIds) {
    added.delete(sheetId)
    finalNames.delete(sheetId)
    hidden.delete(sheetId)
    removed.delete(sheetId)
  }

  const plannerNames = new Map(originalNames)
  for (const [sheetId, addition] of added) plannerNames.set(sheetId, addition.name)

  const renames = workbook.sheets.flatMap((sheet) => {
    if (removed.has(sheet.id)) return []
    const finalName = requiredSheetName(finalNames, sheet.id)
    return finalName === sheet.name ? [] : [{ sheetName: sheet.name, newName: finalName }]
  })

  const additions = [...added.entries()]
    .sort((left, right) => left[1].sequence - right[1].sequence)
    .map(([, addition]) => ({
      name: addition.name,
      ...(addition.sourceSheetId === undefined
        ? {}
        : { sourceSheetName: requiredSheetName(originalNames, addition.sourceSheetId) }),
    }))

  const removals = workbook.sheets
    .filter((sheet) => removed.has(sheet.id))
    .map((sheet) => sheet.name)

  const hiddenChanges = [...hidden.entries()]
    .filter(([sheetId]) => !removed.has(sheetId))
    .map(([sheetId, isHidden]) => ({
      sheetName: originalNames.get(sheetId) ?? requiredSheetName(finalNames, sheetId),
      hidden: isHidden,
    }))

  const order = request.sheetOrder.map((sheetId) => {
    if (removed.has(sheetId)) {
      throw new Error(`Removed worksheet ${sheetId} is still present in the final sheet order.`)
    }
    return requiredSheetName(finalNames, sheetId)
  })

  const expectedIds = new Set([
    ...workbook.sheets.filter((sheet) => !removed.has(sheet.id)).map((sheet) => sheet.id),
    ...added.keys(),
  ])
  if (
    order.length !== expectedIds.size ||
    request.sheetOrder.some((sheetId) => !expectedIds.has(sheetId))
  ) {
    throw new Error('Final sheet order does not match the saved worksheet set.')
  }
  if (new Set(request.sheetOrder).size !== request.sheetOrder.length) {
    throw new Error('Final sheet order contains a duplicate worksheet id.')
  }

  return {
    plannerNames,
    plan: {
      renames,
      additions,
      removals,
      order,
      hiddenChanges,
      orderChanged,
    },
  }
}

function toPlannerStructuralOps(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): PlannerStructuralOps {
  const bySheet = new Map<string, PlannerStructuralOp[]>()
  for (const structuralOp of request.structuralOps) {
    const { sheetId, ...op } = structuralOp
    const sheetName = requiredSheetName(names, sheetId)
    const ops = bySheet.get(sheetName) ?? []
    ops.push(op as PlannerStructuralOp)
    bySheet.set(sheetName, ops)
  }
  return [...bySheet].map(([sheetName, ops]) => ({ sheetName, ops }))
}

function toPlannerCellEdits(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): CellEdit[] {
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
  names: ReadonlyMap<string, string>,
): PlannerFilterStates {
  return request.filterStates.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    filter: state.filter,
    hiddenRows: state.hiddenRows,
    visibilityRange: state.visibilityRange,
  }))
}

function toPlannerHyperlinkEdits(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): PlannerHyperlinkEdits {
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
  names: ReadonlyMap<string, string>,
): PlannerCfStates {
  return request.cfStates.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    rules: state.rules,
  }))
}

function toPlannerDvStates(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): PlannerDvStates {
  return request.dvStates.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    rules: state.rules,
  }))
}

function toPlannerSheetProtections(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): PlannerSheetProtections {
  return request.sheetProtections.map((state) => ({
    sheetName: requiredSheetName(names, state.sheetId),
    protected: state.protected,
  }))
}

function toPlannerPageSetupStates(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): PlannerPageSetupStates {
  return request.pageSetupStates.map(({ sheetId, ...state }) => ({
    sheetName: requiredSheetName(names, sheetId),
    ...state,
  }))
}

function toPlannerNoteStates(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): PlannerNoteStates {
  return request.noteStates.map(({ sheetId, notes }) => ({
    sheetName: requiredSheetName(names, sheetId),
    notes,
  }))
}

function toPlannerFormulaValues(
  request: WorkbookSaveRequest,
  names: ReadonlyMap<string, string>,
): SheetFormulaValues[] {
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

  const sheetContext = buildSheetPlanContext(request, workbook)
  const manifest = await getXlsxArchiveManifest(request.sessionId)
  const source = createEngineEntrySource(request.sessionId, manifest)
  const plan = await planCellEditsToXlsx(
    source,
    toPlannerCellEdits(request, sheetContext.plannerNames),
    toPlannerStructuralOps(request, sheetContext.plannerNames),
    request.chartEdits,
    sheetContext.plan,
    toPlannerFilterStates(request, sheetContext.plannerNames),
    toPlannerHyperlinkEdits(request, sheetContext.plannerNames),
    toPlannerCfStates(request, sheetContext.plannerNames),
    toPlannerDvStates(request, sheetContext.plannerNames),
    toPlannerSheetProtections(request, sheetContext.plannerNames),
    request.definedNamesState,
    [],
    toPlannerPageSetupStates(request, sheetContext.plannerNames),
    toPlannerNoteStates(request, sheetContext.plannerNames),
    [],
    [],
    [],
    [],
    [],
    [],
    toPlannerFormulaValues(request, sheetContext.plannerNames),
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
