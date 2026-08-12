import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

// Valid 1x1 PNG; the Host picker and XLSX save path must preserve the exact bytes.
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WlZ4xkAAAAASUVORK5CYII='

async function createImageWorkbook(path: string): Promise<void> {
  const zip = new JSZip()
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Image host</t></is></c></row></sheetData>
</worksheet>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

const emptySaveRequest = {
  edits: [],
  structuralOps: [],
  chartEdits: [],
  visualEdits: [],
  visualAdditions: [],
  tableAdditions: [],
  pivotAdditions: [],
  sheetOps: [],
  sheetOrder: [],
  filterStates: [],
  hyperlinkEdits: [],
  cfStates: [],
  dvStates: [],
  pageSetupStates: [],
  noteStates: [],
  formulaValues: [],
  pivotCacheRefreshPaths: [],
  pivotRefreshUpdates: [],
  sheetProtections: [],
  sparklineAdditions: [],
  definedNamesState: null,
}

test.describe('Sheets Web Host file integration', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('opens a workbook selected through the Host token/read-file flow', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-host-open-'))
    const workbookPath = join(directory, 'uc-webos-open.xlsx')
    await createImageWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

    const fileChooserPromise = page.waitForEvent('filechooser')
    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=uc-webos-open.xlsx'),
    )
    const selectedWorkbookPromise = editorFrame.locator('body').evaluate(async () => {
      const api = (window as typeof window & { desktopApi: any }).desktopApi
      return api.selectWorkbook()
    })
    const fileChooser = await fileChooserPromise
    await fileChooser.setFiles(workbookPath)

    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      name: string
      sheets: Array<{ id: string; name: string }>
    }
    const selectedWorkbook = await selectedWorkbookPromise

    expect(opened.name).toBe('uc-webos-open.xlsx')
    expect(opened.sheets[0]?.name).toBe('Data')
    expect(selectedWorkbook.sessionId).toBe(opened.sessionId)
    expect(selectedWorkbook.name).toBe('uc-webos-open.xlsx')
    expect(selectedWorkbook.sheets[0]?.name).toBe('Data')
  })

  test('picks an image through the Host, inserts it, then moves the saved drawing anchor', async ({
    page,
  }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-image-'))
    const workbookPath = join(directory, 'web-excel-image.xlsx')
    await createImageWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=web-excel-image.xlsx'),
    )
    await page.locator('#xlsx-picker').setInputFiles(workbookPath)
    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      sheets: Array<{ id: string; name: string }>
    }
    const sheetId = opened.sheets[0]!.id

    // Exercise the same office:pick-file bridge that UC/Web OS will implement.
    const fileChooserPromise = page.waitForEvent('filechooser')
    const imageResultPromise = editorFrame.locator('body').evaluate(async () => {
      const api = (window as typeof window & { desktopApi: any }).desktopApi
      return api.readLocalImage({ path: 'host-picker://insert-image' })
    })
    const fileChooser = await fileChooserPromise
    await fileChooser.setFiles({
      name: 'uc-webos-image.png',
      mimeType: 'image/png',
      buffer: Buffer.from(PNG_BASE64, 'base64'),
    })
    const pickedImage = await imageResultPromise
    expect(pickedImage).toEqual({ mediaType: 'image/png', base64: PNG_BASE64 })

    const inserted = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          ...payload.empty,
          sessionId: payload.sessionId,
          mode: 'save',
          visualAdditions: [
            {
              sheetId: payload.sheetId,
              anchor: {
                fromRow: 1,
                fromColumn: 2,
                fromRowOffset: 0,
                fromColumnOffset: 0,
                toRow: 8,
                toColumn: 6,
                toRowOffset: 0,
                toColumnOffset: 0,
              },
              image: payload.image,
            },
          ],
        })
      },
      {
        sessionId: opened.sessionId,
        sheetId,
        image: pickedImage,
        empty: emptySaveRequest,
      },
    )

    expect(inserted.canceled).toBe(false)
    expect(inserted.touchedEntries).toContain('xl/media/image1.png')
    expect(inserted.touchedEntries).toContain('xl/drawings/drawing1.xml')
    expect(inserted.file.visuals?.some((visual: any) => visual.kind === 'image')).toBe(true)

    const moved = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          ...payload.empty,
          sessionId: payload.sessionId,
          mode: 'save',
          visualEdits: [
            {
              drawingPath: 'xl/drawings/drawing1.xml',
              drawingIndex: 0,
              anchor: {
                fromRow: 3,
                fromColumn: 4,
                fromRowOffset: 0,
                fromColumnOffset: 0,
                toRow: 10,
                toColumn: 8,
                toRowOffset: 0,
                toColumnOffset: 0,
              },
            },
          ],
        })
      },
      { sessionId: inserted.file.sessionId, empty: emptySaveRequest },
    )

    expect(moved.canceled).toBe(false)
    expect(moved.touchedEntries).toContain('xl/drawings/drawing1.xml')
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#download-button')).toBeEnabled()

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('web-excel-image.xlsx')
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const downloadedZip = await JSZip.loadAsync(await readFile(downloadPath!))
    const media = await downloadedZip.file('xl/media/image1.png')?.async('uint8array')
    expect(media).toBeDefined()
    expect(Array.from(media!.slice(0, 8))).toEqual([137, 80, 78, 71, 13, 10, 26, 10])

    const worksheetXml = await downloadedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(worksheetXml).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )
    expect(worksheetXml).toContain('<drawing r:id="rId1"/>')

    const sheetRels = await downloadedZip
      .file('xl/worksheets/_rels/sheet1.xml.rels')
      ?.async('text')
    expect(sheetRels).toContain('relationships/drawing')
    expect(sheetRels).toContain('Target="../drawings/drawing1.xml"')

    const drawingXml = await downloadedZip.file('xl/drawings/drawing1.xml')?.async('text')
    expect(drawingXml).toContain('<xdr:pic>')
    expect(drawingXml).toContain('<xdr:col>4</xdr:col>')
    expect(drawingXml).toContain('<xdr:row>3</xdr:row>')
    expect(drawingXml).toContain('<xdr:col>8</xdr:col>')
    expect(drawingXml).toContain('<xdr:row>10</xdr:row>')
    expect(drawingXml).toContain('r:embed="rId1"')

    const drawingRels = await downloadedZip.file('xl/drawings/_rels/drawing1.xml.rels')?.async('text')
    expect(drawingRels).toContain('relationships/image')
    expect(drawingRels).toContain('Target="../media/image1.png"')

    const contentTypes = await downloadedZip.file('[Content_Types].xml')?.async('text')
    expect(contentTypes).toContain('Extension="png"')
    expect(contentTypes).toContain('ContentType="image/png"')
    expect(contentTypes).toContain('PartName="/xl/drawings/drawing1.xml"')
  })
})
