import type { WorkbookSaveRequest } from '../shared/desktop-api'
import {
  planCellEditsToXlsx,
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

/**
 * Browser preservation-save path. The same mutation planner used by Electron
 * reads package parts through the Rust session API; Rust only reassembles the
 * planned replacement/add/remove sets and returns standard XLSX bytes.
 */
export async function saveWorkbookRequestViaEngine(
  request: WorkbookSaveRequest,
  name: string,
): Promise<SavedXlsxWorkbook & { touchedEntries: readonly string[] }> {
  if (request.sheetOps.length > 0) {
    throw new Error('Sheet management saves are not enabled in Sheets Web yet.')
  }

  const manifest = await getXlsxArchiveManifest(request.sessionId)
  const source = createEngineEntrySource(request.sessionId, manifest)
  const plan = await planCellEditsToXlsx(
    source,
    request.edits,
    request.structuralOps,
    request.chartEdits,
    undefined,
    request.filterStates,
    request.hyperlinkEdits,
    request.cfStates,
    request.dvStates,
    request.sheetProtections,
    request.definedNamesState,
    request.visualAdditions,
    request.pageSetupStates,
    request.noteStates,
    request.tableAdditions,
    request.pivotAdditions,
    request.pivotCacheRefreshPaths,
    request.pivotRefreshUpdates,
    request.visualEdits,
    request.sparklineAdditions,
    request.formulaValues,
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
