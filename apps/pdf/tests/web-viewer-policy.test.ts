import { describe, expect, it } from 'vitest'
import { isPdfOfficeFile, PDF_WEB_ACCEPT, PDF_WEB_FEATURES } from '../src/web/viewer-policy'

describe('PDF Web viewer-only policy', () => {
  it('keeps all mutation capabilities disabled', () => {
    expect(PDF_WEB_FEATURES.view).toBe(true)
    expect(PDF_WEB_FEATURES.open).toBe(true)
    expect(PDF_WEB_FEATURES.search).toBe(true)
    expect(PDF_WEB_FEATURES.edit).toBe(false)
    expect(PDF_WEB_FEATURES.save).toBe(false)
    expect(PDF_WEB_FEATURES.annotations).toBe(false)
    expect(PDF_WEB_FEATURES.forms).toBe(false)
    expect(PDF_WEB_FEATURES.signatures).toBe(false)
    expect(PDF_WEB_FEATURES.stamps).toBe(false)
    expect(PDF_WEB_FEATURES.imageEditing).toBe(false)
    expect(PDF_WEB_FEATURES.aiEditing).toBe(false)
  })

  it('accepts PDFs by MIME type or file extension', () => {
    expect(PDF_WEB_ACCEPT).toEqual(['application/pdf', '.pdf'])
    expect(isPdfOfficeFile({ name: 'report.bin', mimeType: 'application/pdf' })).toBe(true)
    expect(isPdfOfficeFile({ name: 'REPORT.PDF', mimeType: 'application/octet-stream' })).toBe(true)
    expect(isPdfOfficeFile({ name: 'report.docx', mimeType: 'application/octet-stream' })).toBe(false)
  })
})
