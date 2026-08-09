import {
  DEFAULT_EMBEDDED_OFFICE_CAPABILITIES,
  DEFAULT_STANDALONE_OFFICE_CAPABILITIES,
  type OfficeHostApi,
  type OfficeHostCapabilities,
} from '@genoffice/office-host-api'
import type { EditorIframeBridge, WebRuntimeMode } from '@genoffice/web-runtime'
import type { DocsWebDesktopController } from './desktop-api'

const CAPABILITY_CLASSES = [
  'office-web',
  'office-ai-enabled',
  'office-autosave-editor',
  'office-page-crop-marks',
  'office-can-open',
  'office-can-save',
  'office-can-save-as',
  'office-can-save-history',
  'office-can-export-docx',
  'office-can-close',
] as const

function applyCapabilityClasses(capabilities: OfficeHostCapabilities): void {
  const root = document.documentElement
  root.classList.add('office-web')
  root.classList.toggle('office-ai-enabled', capabilities.ai)
  root.classList.toggle('office-autosave-editor', capabilities.autoSave === 'editor')
  root.classList.toggle('office-page-crop-marks', capabilities.pageCropMarks)
  root.classList.toggle('office-can-open', capabilities.open)
  root.classList.toggle('office-can-save', capabilities.save)
  root.classList.toggle('office-can-save-as', capabilities.saveAs)
  root.classList.toggle('office-can-save-history', capabilities.saveHistoryVersion)
  root.classList.toggle('office-can-export-docx', capabilities.exportDocx)
  root.classList.toggle('office-can-close', capabilities.close)
}

export interface WebHostPolicyController {
  getCapabilities(): OfficeHostCapabilities
  destroy(): void
}

export function installWebHostPolicy(
  mode: WebRuntimeMode,
  bridge?: EditorIframeBridge,
): WebHostPolicyController {
  let capabilities: OfficeHostCapabilities = {
    ...(mode === 'embedded'
      ? DEFAULT_EMBEDDED_OFFICE_CAPABILITIES
      : DEFAULT_STANDALONE_OFFICE_CAPABILITIES),
  }

  const apply = (patch?: Partial<OfficeHostCapabilities>) => {
    if (patch) capabilities = { ...capabilities, ...patch }
    applyCapabilityClasses(capabilities)
  }
  apply()

  const unsubscribe = bridge?.subscribe((message) => {
    if (
      (message.type === 'office:init' || message.type === 'office:new') &&
      message.payload.capabilities
    ) {
      apply(message.payload.capabilities)
    }
  })

  return {
    getCapabilities: () => ({ ...capabilities }),
    destroy: () => {
      unsubscribe?.()
      document.documentElement.classList.remove(...CAPABILITY_CLASSES)
    },
  }
}

/**
 * Existing renderer code already distinguishes saveDocx() and saveDocxAs().
 * This adapter keeps that stable desktop-facing API while attaching explicit
 * save/saveAs intent to the browser host call. No renderer/file-action fork is required.
 */
export function installWebSaveModeAdapter(
  controller: DocsWebDesktopController,
  host: OfficeHostApi,
): () => void {
  const originalHostSave = host.saveDocument.bind(host)
  const originalSaveAs = controller.desktopApi.saveDocxAs.bind(controller.desktopApi)
  let nextSaveAs = false

  host.saveDocument = async (input) => {
    const mode = input.mode ?? (nextSaveAs ? 'saveAs' : 'save')
    nextSaveAs = false
    return originalHostSave({ ...input, mode })
  }

  controller.desktopApi.saveDocxAs = async (defaultName, data) => {
    nextSaveAs = true
    try {
      return await originalSaveAs(defaultName, data)
    } finally {
      nextSaveAs = false
    }
  }

  return () => {
    host.saveDocument = originalHostSave
    controller.desktopApi.saveDocxAs = originalSaveAs
  }
}
