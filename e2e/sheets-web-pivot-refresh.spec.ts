import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

const PIVOT_CACHE_PATH = 'xl/pivotCache/pivotCacheDefinition1.xml'

async function createPivotRefreshWorkbook(path: string, conflict: boolean): Promise<void> {
  const zip = new JSZip()
  const conflictCell = conflict ? '<c r="H2" t="inlineStr"><is><t>KEEP</t></is></c>' : ''
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <pivotCaches><pivotCache cacheId="1" r:id="rId3"/></pivotCaches>
</workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:${conflict ? 'H4' : 'G4'}"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c>
      <c r="F1" t="inlineStr"><is><t>Category</t></is></c><c r="G1" t="inlineStr"><is><t>Sum of Value</t></is></c>
    </row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>10</v></c><c r="F2" t="inlineStr"><is><t>A</t></is></c><c r="G2"><v>10</v></c>${conflictCell}</row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>B</t></is></c><c r="B3"><v>20</v></c><c r="F3" t="inlineStr"><is><t>B</t></is></c><c r="G3"><v>20</v></c></row>
    <row r="4"><c r="F4" t="inlineStr"><is><t>Grand Total</t></is></c><c r="G4"><v>30</v></c></row>
  </sheetData>
  <pivotTableParts count="1"><pivotTablePart r:id="rId1"/></pivotTableParts>
</worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
    'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Pivot1" cacheId="1" dataCaption="Values">
  <location ref="F1:G4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2"><pivotField axis="axisRow" showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <rowItems count="3"><i><x v="0"/></i><i><x v="1"/></i><i t="grand"><x v="0"/></i></rowItems>
  <dataFields count="1"><dataField name="Sum of Value" fld="1" subtotal="sum"/></dataFields>
</pivotTableDefinition>`,
    'xl/pivotTables/_rels/pivotTable1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`,
    [PIVOT_CACHE_PATH]: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="2">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Data"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="Category"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField>
    <cacheField name="Value"><sharedItems containsNumber="1" minValue="10" maxValue="20"/></cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
    'xl/pivotCache/pivotCacheRecords1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2"><r><x v="0"/><n v="10"/></r><r><x v="1"/><n v="20"/></r></pivotCacheRecords>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

function emptySaveRequest(sessionId: string) {
  return {
    sessionId,
    mode: 'save' as const,
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
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
  }
}

async function openWorkbook(page: import('@playwright/test').Page, workbookPath: string, name: string) {
  await page.goto(hostUrl!)
  const editorFrame = page.frameLocator('#office-frame')
  await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

  const openResponsePromise = page.waitForResponse(
    (response) =>
      response.request().method() === 'POST' &&
      response.url().includes(`/xlsx-engine/v1/workbooks?name=${name}`),
  )
  await page.locator('#xlsx-picker').setInputFiles(workbookPath)
  const opened = (await (await openResponsePromise).json()) as {
    sessionId: string
    sheets: Array<{ id: string; name: string }>
  }
  expect(opened.sheets[0]?.name).toBe('Data')
  return { editorFrame, opened }
}

test.describe('Sheets Web existing Pivot refresh', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('expands an existing Pivot output area and marks its cache refresh-on-load', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')
    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-pivot-refresh-'))
    const workbookPath = join(directory, 'web-excel-pivot-refresh.xlsx')
    await createPivotRefreshWorkbook(workbookPath, false)

    const { editorFrame, opened } = await openWorkbook(
      page,
      workbookPath,
      'web-excel-pivot-refresh.xlsx',
    )
    const sheetId = opened.sheets[0]!.id
    const saved = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          ...payload.empty,
          pivotCacheRefreshPaths: [payload.cachePath],
          pivotRefreshUpdates: [
            {
              cachePath: payload.cachePath,
              sheetId: payload.sheetId,
              newOutputRef: 'F1:H4',
            },
          ],
        })
      },
      {
        empty: emptySaveRequest(opened.sessionId),
        cachePath: PIVOT_CACHE_PATH,
        sheetId,
      },
    )

    expect(saved.canceled).toBe(false)
    expect(saved.touchedEntries).toEqual(
      expect.arrayContaining(['xl/pivotTables/pivotTable1.xml', PIVOT_CACHE_PATH]),
    )
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const zip = await JSZip.loadAsync(await readFile(downloadPath!))
    const pivotXml = await zip.file('xl/pivotTables/pivotTable1.xml')?.async('text')
    const cacheXml = await zip.file(PIVOT_CACHE_PATH)?.async('text')
    expect(pivotXml).toMatch(/<location\b[^>]*\bref="F1:H4"/)
    expect(cacheXml).toMatch(/<pivotCacheDefinition\b[^>]*\brefreshOnLoad="1"/)
  })

  test('fails closed when the newly occupied Pivot area contains ordinary worksheet data', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')
    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-pivot-conflict-'))
    const workbookPath = join(directory, 'web-excel-pivot-conflict.xlsx')
    await createPivotRefreshWorkbook(workbookPath, true)

    const { editorFrame, opened } = await openWorkbook(
      page,
      workbookPath,
      'web-excel-pivot-conflict.xlsx',
    )
    const sheetId = opened.sheets[0]!.id
    const error = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        try {
          await api.saveWorkbookEdits({
            ...payload.empty,
            pivotCacheRefreshPaths: [payload.cachePath],
            pivotRefreshUpdates: [
              {
                cachePath: payload.cachePath,
                sheetId: payload.sheetId,
                newOutputRef: 'F1:H4',
              },
            ],
          })
          return null
        } catch (cause) {
          return cause instanceof Error ? cause.message : String(cause)
        }
      },
      {
        empty: emptySaveRequest(opened.sessionId),
        cachePath: PIVOT_CACHE_PATH,
        sheetId,
      },
    )

    expect(error).toContain('conflicts with existing worksheet content')

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const zip = await JSZip.loadAsync(await readFile(downloadPath!))
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    const pivotXml = await zip.file('xl/pivotTables/pivotTable1.xml')?.async('text')
    const cacheXml = await zip.file(PIVOT_CACHE_PATH)?.async('text')
    expect(sheetXml).toContain('<t>KEEP</t>')
    expect(pivotXml).toMatch(/<location\b[^>]*\bref="F1:G4"/)
    expect(cacheXml).not.toMatch(/<pivotCacheDefinition\b[^>]*\brefreshOnLoad="1"/)
  })
})
