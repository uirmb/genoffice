import {
  DEFAULT_EMBEDDED_OFFICE_CAPABILITIES,
  DEFAULT_STANDALONE_OFFICE_CAPABILITIES,
  type OfficeHostCapabilities,
} from '@genoffice/office-host-api'
import type { EditorIframeBridge, WebRuntimeMode } from '@genoffice/web-runtime'

const CAPABILITY_CLASSES = [
  'office-web',
  'office-ai-enabled',
  'office-autosave-editor',
  'office-can-open',
  'office-can-save',
  'office-can-save-as',
  'office-can-save-history',
  'office-can-export-pptx',
  'office-can-close',
] as const

function applyClasses(capabilities: OfficeHostCapabilities): void {
  const root = document.documentElement
  root.classList.add('office-web')
  root.classList.toggle('office-ai-enabled', capabilities.ai)
  root.classList.toggle('office-autosave-editor', capabilities.autoSave === 'editor')
  root.classList.toggle('office-can-open', capabilities.open)
  root.classList.toggle('office-can-save', capabilities.save)
  root.classList.toggle('office-can-save-as', capabilities.saveAs)
  root.classList.toggle('office-can-save-history', capabilities.saveHistoryVersion)
  root.classList.toggle('office-can-export-pptx', capabilities.exportPptx)
  root.classList.toggle('office-can-close', capabilities.close)
}

export interface SlidesWebHostPolicy {
  getCapabilities(): OfficeHostCapabilities
  destroy(): void
}

export function installSlidesWebHostPolicy(
  mode: WebRuntimeMode,
  bridge?: EditorIframeBridge,
): SlidesWebHostPolicy {
  let capabilities: OfficeHostCapabilities = {
    ...(mode === 'embedded'
      ? DEFAULT_EMBEDDED_OFFICE_CAPABILITIES
      : DEFAULT_STANDALONE_OFFICE_CAPABILITIES),
  }

  const apply = (patch?: Partial<OfficeHostCapabilities>) => {
    if (patch) capabilities = { ...capabilities, ...patch }
    applyClasses(capabilities)
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
