import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

async function createDefinedNamesWorkbook(path: string): Promise<void> {
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
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName>
    <definedName name="OldName">Data!$A$1</definedName>
  </definedNames>
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
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c></row>
    <row r="2"><c r="A2"><v>30</v></c><c r="B2"><v>40</v></c></row>
  </sheetData>
</worksheet>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  await writeFile(
    path,
    await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
  )
}

test.describe('Sheets Web defined names', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('rewrites modeled names while preserving Excel built-ins', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-names-'))
    const workbookPath = join(directory, 'web-excel-defined-names.xlsx')
    await createDefinedNamesWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=web-excel-defined-names.xlsx'),
    )
    await page.locator('#xlsx-picker').setInputFiles(workbookPath)
    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      sheets: Array<{ id: string; name: string }>
    }
    expect(opened.sheets[0]?.name).toBe('Data')

    const saved = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          sessionId: payload.sessionId,
          mode: 'save',
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
          definedNamesState: {
            names: [
              { name: 'Revenue', formula: '=Data!$B$1' },
              { name: 'LocalTotal', formula: '=Data!$A$1', sheetIndex: 0 },
            ],
            preserveNames: [],
          },
        })
      },
      { sessionId: opened.sessionId },
    )

    expect(saved.canceled).toBe(false)
    expect(saved.touchedEntries).toContain('xl/workbook.xml')
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#download-button')).toBeEnabled()

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('web-excel-defined-names.xlsx')
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const downloadedZip = await JSZip.loadAsync(await readFile(downloadPath!))
    const workbookXml = await downloadedZip.file('xl/workbook.xml')?.async('text')
    expect(workbookXml).toContain(
      '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName>',
    )
    expect(workbookXml).not.toContain('name="OldName"')
    expect(workbookXml).toContain('<definedName name="Revenue">Data!$B$1</definedName>')
    expect(workbookXml).toContain(
      '<definedName name="LocalTotal" localSheetId="0">Data!$A$1</definedName>',
    )
  })
})
