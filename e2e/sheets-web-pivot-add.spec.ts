import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

async function createPivotSourceWorkbook(path: string): Promise<void> {
  const zip = new JSZip()
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Pivot" sheetId="2" r:id="rId2"/></sheets>
</workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
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
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>B</t></is></c><c r="B3"><v>20</v></c></row>
  </sheetData>
</worksheet>`,
    'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

test.describe('Sheets Web pivot creation', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('creates a populated native pivot package and reads it from the saved session', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-pivot-add-'))
    const workbookPath = join(directory, 'web-excel-pivot-add.xlsx')
    await createPivotSourceWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=web-excel-pivot-add.xlsx'),
    )
    await page.locator('#xlsx-picker').setInputFiles(workbookPath)
    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      sheets: Array<{ id: string; name: string }>
    }
    expect(opened.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Pivot'])
    const sourceSheetId = opened.sheets.find((sheet) => sheet.name === 'Data')!.id
    const pivotSheetId = opened.sheets.find((sheet) => sheet.name === 'Pivot')!.id

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
          pivotAdditions: [
            {
              sheetId: payload.pivotSheetId,
              sourceSheetId: payload.sourceSheetId,
              sourceArea: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
              location: { startRow: 0, endRow: 3, startColumn: 5, endColumn: 6 },
              name: 'WebPivot',
              fieldNames: ['Category', 'Value'],
              rowFieldIndices: [0],
              rowItems: ['A', 'B'],
              rowLevelItems: [['A', 'B']],
              rowLines: [
                { t: 'data', members: [0] },
                { t: 'data', members: [1] },
              ],
              values: [{ fieldIndex: 1, agg: 'sum', name: 'Sum of Value' }],
            },
          ],
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
      {
        sessionId: opened.sessionId,
        sourceSheetId,
        pivotSheetId,
      },
    )

    expect(saved.canceled).toBe(false)
    for (const path of [
      'xl/pivotTables/pivotTable1.xml',
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'xl/pivotCache/pivotCacheRecords1.xml',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      '[Content_Types].xml',
    ]) {
      expect(saved.touchedEntries).toContain(path)
    }

    const definition = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readPivotDefinition({
          sessionId: payload.sessionId,
          path: 'xl/pivotTables/pivotTable1.xml',
          cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
        })
      },
      { sessionId: saved.file.sessionId },
    )
    expect(definition.outputRef).toBe('F1:G4')
    expect(definition.sourceSheet).toBe('Data')
    expect(definition.sourceRef).toBe('A1:B3')
    expect(definition.fields).toEqual([
      { name: 'Category', sharedItems: ['A', 'B'] },
      { name: 'Value', sharedItems: [] },
    ])
    expect(definition.rowFields).toEqual([0])
    expect(definition.dataFields).toEqual([
      { name: 'Sum of Value', field: 1, subtotal: 'sum' },
    ])
    expect(definition.unsupported).toEqual([])

    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const downloadedZip = await JSZip.loadAsync(await readFile(downloadPath!))
    const recordsXml = await downloadedZip
      .file('xl/pivotCache/pivotCacheRecords1.xml')
      ?.async('text')
    expect(recordsXml).toContain('count="2"')
    expect(recordsXml).toContain('<r><x v="0"/><n v="10"/></r>')
    expect(recordsXml).toContain('<r><x v="1"/><n v="20"/></r>')

    const cacheXml = await downloadedZip
      .file('xl/pivotCache/pivotCacheDefinition1.xml')
      ?.async('text')
    expect(cacheXml).toContain('recordCount="2"')
    expect(cacheXml).not.toContain('refreshOnLoad="1"')
    expect(cacheXml).toContain('<worksheetSource ref="A1:B3" sheet="Data"/>')
    expect(cacheXml).toContain(
      '<cacheField name="Category" numFmtId="0"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems>',
    )

    const pivotXml = await downloadedZip.file('xl/pivotTables/pivotTable1.xml')?.async('text')
    expect(pivotXml).toContain('name="WebPivot" cacheId="1"')
    expect(pivotXml).toContain('<location ref="F1:G4"')
    expect(pivotXml).toContain('<rowFields count="1"><field x="0"/></rowFields>')
    expect(pivotXml).toContain('<dataField name="Sum of Value" fld="1"')

    const pivotRels = await downloadedZip
      .file('xl/pivotTables/_rels/pivotTable1.xml.rels')
      ?.async('text')
    expect(pivotRels).toContain('pivotCacheDefinition')
    expect(pivotRels).toContain('Target="../pivotCache/pivotCacheDefinition1.xml"')

    const cacheRels = await downloadedZip
      .file('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')
      ?.async('text')
    expect(cacheRels).toContain('pivotCacheRecords')
    expect(cacheRels).toContain('Target="pivotCacheRecords1.xml"')

    const pivotSheetRels = await downloadedZip
      .file('xl/worksheets/_rels/sheet2.xml.rels')
      ?.async('text')
    expect(pivotSheetRels).toContain('relationships/pivotTable')
    expect(pivotSheetRels).toContain('Target="../pivotTables/pivotTable1.xml"')

    const workbookXml = await downloadedZip.file('xl/workbook.xml')?.async('text')
    expect(workbookXml).toMatch(
      /<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/,
    )
    const workbookRels = await downloadedZip.file('xl/_rels/workbook.xml.rels')?.async('text')
    expect(workbookRels).toContain('pivotCacheDefinition')
    expect(workbookRels).toContain('Target="pivotCache/pivotCacheDefinition1.xml"')

    const contentTypes = await downloadedZip.file('[Content_Types].xml')?.async('text')
    expect(contentTypes).toContain('PartName="/xl/pivotTables/pivotTable1.xml"')
    expect(contentTypes).toContain('pivotCacheDefinition+xml')
    expect(contentTypes).toContain('pivotCacheRecords+xml')
  })
})
