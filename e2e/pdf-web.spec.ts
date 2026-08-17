import { expect, test } from '@playwright/test'
import { PDFDocument, StandardFonts } from 'pdf-lib'

const PDF_WEB_URL = process.env.PDF_WEB_E2E_URL ?? 'http://127.0.0.1:5276/'
const PDF_WEB_HOST_URL = process.env.PDF_WEB_HOST_E2E_URL ?? 'http://127.0.0.1:8083/'

async function makePdf(pages: string[]): Promise<Buffer> {
  const document = await PDFDocument.create()
  const font = await document.embedFont(StandardFonts.Helvetica)
  for (const [index, text] of pages.entries()) {
    const page = document.addPage([595, 842])
    page.drawText(`GenOffice PDF Web ${index + 1}`, { x: 64, y: 760, size: 22, font })
    page.drawText(text, { x: 64, y: 700, size: 14, font })
  }
  return Buffer.from(await document.save())
}

test('standalone PDF Web exposes a viewer-only empty state', async ({ page }) => {
  await page.goto(PDF_WEB_URL)

  await expect(page.getByRole('heading', { name: 'PDF 预览' })).toBeVisible()
  await expect(
    page.getByText('Web 版本仅提供阅读能力，不包含编辑、批注、签名或表单修改。'),
  ).toBeVisible()
  await expect(page.getByRole('button', { name: '打开 PDF' })).toBeVisible()
  await expect(page.getByText('只读')).toBeVisible()
  await expect(page.getByRole('button', { name: /保存|另存为|签名|批注/ })).toHaveCount(0)
})

test('embedded PDF Web opens, searches, and transactionally replaces PDFs', async ({ page }) => {
  const firstPdf = await makePdf(['First page', 'Search Marker 2026'])
  const secondPdf = await makePdf(['Replacement document'])

  await page.goto(PDF_WEB_HOST_URL)
  const officeFrame = page.frameLocator('#office-frame')
  await expect(officeFrame.getByRole('heading', { name: 'PDF 预览' })).toBeVisible()

  await page.locator('#pdf-picker').setInputFiles({
    name: 'viewer-two-pages.pdf',
    mimeType: 'application/pdf',
    buffer: firstPdf,
  })

  await expect(officeFrame.locator('.pdf-web-page')).toHaveCount(2)
  await expect(officeFrame.getByText('2 页')).toBeVisible()
  await expect(officeFrame.getByText('viewer-two-pages.pdf')).toBeVisible()
  await expect(page.locator('#host-state')).toHaveText('opened')

  const search = officeFrame.getByRole('textbox', { name: '搜索文档' })
  await search.fill('Search Marker 2026')
  await officeFrame.getByRole('button', { name: '查找' }).click()
  await expect(officeFrame.locator('.pdf-web-search-count')).toHaveText('1/1')
  await expect(officeFrame.locator('#pdf-web-page-2')).toHaveClass(/search-target/)

  const chooserPromise = page.waitForEvent('filechooser')
  await officeFrame.getByRole('button', { name: '打开' }).click()
  const chooser = await chooserPromise
  await chooser.setFiles({
    name: 'viewer-replacement.pdf',
    mimeType: 'application/pdf',
    buffer: secondPdf,
  })

  await expect(officeFrame.locator('.pdf-web-page')).toHaveCount(1)
  await expect(officeFrame.getByText('1 页')).toBeVisible()
  await expect(officeFrame.getByText('viewer-replacement.pdf')).toBeVisible()
  await expect(page.locator('#file-name')).toHaveText('viewer-replacement.pdf')
  await expect(page.locator('#host-state')).toHaveText('opened')
})
