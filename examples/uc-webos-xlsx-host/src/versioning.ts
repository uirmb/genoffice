export interface VersionedFileAccess {
  version?: string | number | null | undefined
  fileVersion?: string | number | null | undefined
}

function normalizeVersion(value: unknown): string | null {
  if (value === undefined || value === null) return null
  const text = String(value).trim()
  return text || null
}

export function platformAccessVersion(access: VersionedFileAccess): string | null {
  return normalizeVersion(access.version) ?? normalizeVersion(access.fileVersion)
}

export interface VersionConflict {
  code: 'VERSION_CONFLICT'
  expectedVersion: string
  actualVersion: string
  error: string
}

/**
 * Compatibility helper for the legacy UC bridge. The stable UC protocol moves
 * optimistic version enforcement into uc.fs.saveCurrentFile so the Bridge no
 * longer owns this check. Keep this helper tolerant of opaque string/number
 * versions until that migration is complete.
 */
export function detectSelectedFileVersionConflict(
  baseVersion: string | number | null | undefined,
  access: VersionedFileAccess,
): VersionConflict | null {
  const expectedVersion = normalizeVersion(baseVersion)
  const actualVersion = platformAccessVersion(access)
  if (!expectedVersion || !actualVersion || expectedVersion === actualVersion) return null

  return {
    code: 'VERSION_CONFLICT',
    expectedVersion,
    actualVersion,
    error: `文件版本已变化（编辑基线 ${expectedVersion}，当前 ${actualVersion}），请重新打开后再保存。`,
  }
}
