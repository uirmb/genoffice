import type { WorkbookPivotDefinition, WorkbookPivotRequest } from '../shared/desktop-api'
import { parsePivotDefinition } from '../gateway/xlsx-pivot'
import { readXlsxArchiveEntries } from './engine-client'

const decoder = new TextDecoder()

/**
 * Browser Pivot reader. The request already carries the exact pivotTable and
 * pivotCacheDefinition part paths discovered from workbook metadata; the Web
 * adapter only reads those session-scoped archive entries and reuses the same
 * fail-closed parser as Electron.
 */
export async function readPivotDefinitionViaEngine(
  request: WorkbookPivotRequest,
): Promise<WorkbookPivotDefinition> {
  const entries = await readXlsxArchiveEntries(request.sessionId, [request.path, request.cachePath])
  const pivotBytes = entries.get(request.path)
  const cacheBytes = entries.get(request.cachePath)
  if (!pivotBytes) throw new Error(`XLSX Engine did not return ${request.path}.`)
  if (!cacheBytes) throw new Error(`XLSX Engine did not return ${request.cachePath}.`)

  return parsePivotDefinition(
    decoder.decode(pivotBytes),
    decoder.decode(cacheBytes),
  ) as WorkbookPivotDefinition
}
