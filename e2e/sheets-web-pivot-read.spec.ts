import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

async function createPivotWorkbook(path: string): Promise<void> {
  const zip = new JSZip()
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
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>B</t></is></c><c r="B3"><v>20</v></c></row>
  </sheetData>
</worksheet>`,
    'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1">
  <location ref="D1:E3" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow" showAll="0"><items count="2"><item x="0"/><item x="1"/></items></pivotField>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <rowItems count="2"><i><x v="0"/></i><i><x v="1"/></i></rowItems>
  <dataFields count="1"><dataField name="Sum of Value" fld="1" subtotal="sum"/></dataFields>
</pivotTableDefinition>`,
    'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Data"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="Category"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField>
    <cacheField name="Value"><sharedItems containsNumber="1" count="2"><n v="10"/><n v="20"/></sharedItems></cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

test.describe('Sheets Web pivot reads', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('parses pivot table and cache definitions from the active session', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-pivot-'))
    const workbookPath = join(directory, 'web-excel-pivot-read.xlsx')
    await createPivotWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=web-excel-pivot-read.xlsx'),
    )
    await page.locator('#xlsx-picker').setInputFiles(workbookPath)
    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      sheets: Array<{ id: string; name: string }>
    }
    expect(opened.sheets[0]?.name).toBe('Data')

    const definition = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readPivotDefinition({
          sessionId: payload.sessionId,
          path: 'xl/pivotTables/pivotTable1.xml',
          cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
        })
      },
      { sessionId: opened.sessionId },
    )

    expect(definition.outputRef).toBe('D1:E3')
    expect(definition.firstDataRow).toBe(1)
    expect(definition.firstDataCol).toBe(1)
    expect(definition.sourceSheet).toBe('Data')
    expect(definition.sourceRef).toBe('A1:B3')
    expect(definition.fields).toEqual([
      { name: 'Category', sharedItems: ['A', 'B'] },
      { name: 'Value', sharedItems: [10, 20] },
    ])
    expect(definition.rowFields).toEqual([0])
    expect(definition.colFields).toEqual([])
    expect(definition.dataFields).toEqual([
      { name: 'Sum of Value', field: 1, subtotal: 'sum' },
    ])
    expect(definition.fieldItems[0]).toEqual([
      { x: 0, hidden: false },
      { x: 1, hidden: false },
    ])
    expect(definition.unsupported).toEqual([])
  })
})
