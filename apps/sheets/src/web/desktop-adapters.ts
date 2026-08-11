import type { OfficeHostApi } from '@genoffice/office-host-api'
import type { DesktopApi } from '../shared/desktop-api'
import { readLocalImageViaHost } from './local-image'
import { readPivotDefinitionViaEngine } from './pivot-reader'

/**
 * Installs browser-only DesktopApi capabilities that sit outside the generic
 * Sheets Web controller. Keeping them in one explicit adapter layer makes the
 * Electron/Web boundary visible and gives platform hosts one integration point.
 */
export function installSheetsWebDesktopAdapters(api: DesktopApi, host: OfficeHostApi): void {
  api.readPivotDefinition = readPivotDefinitionViaEngine
  api.readLocalImage = (request) => readLocalImageViaHost(host, request)
}
