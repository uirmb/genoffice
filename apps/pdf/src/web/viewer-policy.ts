import type { OfficeFile } from '@genoffice/office-host-api'

export const PDF_MIME_TYPE = 'application/pdf'
export const PDF_WEB_ACCEPT = [PDF_MIME_TYPE, '.pdf'] as const

/**
 * PDF Web phase 1 is deliberately read-only. Keep this policy explicit so
 * editor-only features cannot accidentally leak into the browser product.
 */
export const PDF_WEB_FEATURES = Object.freeze({
  view: true,
  open: true,
  thumbnails: true,
  outline: true,
  textSelection: true,
  links: true,
  search: true,
  pageNavigation: true,
  zoom: true,
  edit: false,
  save: false,
  annotations: false,
  forms: false,
  signatures: false,
  stamps: false,
  imageEditing: false,
  aiEditing: false,
})

export function isPdfOfficeFile(file: Pick<OfficeFile, 'name' | 'mimeType'>): boolean {
  if (file.mimeType.toLowerCase() === PDF_MIME_TYPE) return true
  return file.name.toLowerCase().endsWith('.pdf')
}
