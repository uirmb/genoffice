import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

async function createMinimalWorkbook(path: string): Promise<void> {
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
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
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
  <dimension ref="A1:C1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Browser Original</t></is></c><c r="B1"><v>42</v></c><c r="C1"><f>B1*2</f><v>84</v></c></row></sheetData>
</worksheet>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  await writeFile(path, bytes)
}

test.describe('Sheets Web', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('opens, reads formulas, recalculates, saves, rereads, and downloads a real XLSX through the iframe host', async ({
    page,
  }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-'))
    const workbookPath = join(directory, 'web-excel-browser-e2e.xlsx')
    await createMinimalWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })
    await expect(editorFrame.locator('text=GenOffice Sheets Web failed to start')).toHaveCount(0)

    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=web-excel-browser-e2e.xlsx'),
    )
    await page.locator('#xlsx-picker').setInputFiles(workbookPath)
    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      sheets: Array<{ id: string; name: string }>
    }
    expect(opened.sessionId).toBeTruthy()
    expect(opened.sheets[0]?.name).toBe('Sheet1')

    const sheetId = opened.sheets[0]!.id
    const initialRange = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readWorkbookRange({
          sessionId: payload.sessionId,
          sheetId: payload.sheetId,
          range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
        })
      },
      { sessionId: opened.sessionId, sheetId },
    )
    expect(initialRange.cells.find((cell: any) => cell.row === 0 && cell.column === 0)?.value).toBe(
      'Browser Original',
    )
    expect(initialRange.cells.find((cell: any) => cell.row === 0 && cell.column === 1)?.value).toBe(42)

    const formulas = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readWorkbookFormulas({
          sessionId: payload.sessionId,
          sheetId: payload.sheetId,
        })
      },
      { sessionId: opened.sessionId, sheetId },
    )
    expect(formulas.truncated).toBe(false)
    expect(formulas.cells.find((cell: any) => cell.row === 0 && cell.column === 2)?.formula).toBe(
      'B1*2',
    )

    const recalculated = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.recalcWorkbook({
          sessionId: payload.sessionId,
          edits: [
            {
              sheetId: payload.sheetId,
              row: 0,
              column: 1,
              input: '50',
            },
          ],
          reads: [
            {
              sheetId: payload.sheetId,
              range: { startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 },
            },
          ],
        })
      },
      { sessionId: opened.sessionId, sheetId },
    )
    const recalculatedFormula = recalculated.cells.find(
      (cell: any) => cell.row === 0 && cell.column === 2,
    )
    expect(recalculatedFormula?.number).toBe(100)
    expect(recalculatedFormula?.isFormula).toBe(true)

    const saved = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          sessionId: payload.sessionId,
          mode: 'save',
          edits: [
            {
              sheetId: payload.sheetId,
              row: 0,
              column: 0,
              writeValue: true,
              value: 'Browser Saved',
            },
          ],
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
        })
      },
      { sessionId: opened.sessionId, sheetId },
    )

    expect(saved.canceled).toBe(false)
    expect(saved.touchedEntries).toContain('xl/worksheets/sheet1.xml')
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#file-name')).toHaveText('web-excel-browser-e2e.xlsx')
    await expect(page.locator('#download-button')).toBeEnabled()

    const savedSheetId = saved.file.sheets[0].id
    const savedRange = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readWorkbookRange({
          sessionId: payload.sessionId,
          sheetId: payload.sheetId,
          range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        })
      },
      { sessionId: saved.file.sessionId, sheetId: savedSheetId },
    )
    expect(savedRange.cells.find((cell: any) => cell.row === 0 && cell.column === 0)?.value).toBe(
      'Browser Saved',
    )

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('web-excel-browser-e2e.xlsx')
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const downloadedZip = await JSZip.loadAsync(await readFile(downloadPath!))
    const worksheetXml = await downloadedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(worksheetXml).toContain('Browser Saved')
    expect(worksheetXml).toContain('<f>B1*2</f>')
  })
})
