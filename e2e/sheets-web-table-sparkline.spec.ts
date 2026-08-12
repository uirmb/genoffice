import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

async function createTableSparklineWorkbook(path: string): Promise<void> {
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
  <dimension ref="A1:D3"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Beta</t></is></c><c r="B3"><v>20</v></c></row>
  </sheetData>
</worksheet>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

test.describe('Sheets Web table and sparkline saves', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('creates a native table part and x14 sparkline group', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-table-'))
    const workbookPath = join(directory, 'web-excel-table-sparkline.xlsx')
    await createTableSparklineWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=web-excel-table-sparkline.xlsx'),
    )
    await page.locator('#xlsx-picker').setInputFiles(workbookPath)
    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      sheets: Array<{ id: string; name: string }>
    }
    const sheetId = opened.sheets[0]!.id
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
          tableAdditions: [
            {
              sheetId: payload.sheetId,
              area: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
              name: 'WebTable',
              columnNames: ['Name', 'Value'],
              style: 'TableStyleMedium2',
              bandedRows: true,
            },
          ],
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
          sparklineAdditions: [
            {
              sheetId: payload.sheetId,
              type: 'column',
              color: '#336699',
              cells: [{ cell: 'D2', sourceRef: 'Data!B2:B3' }],
            },
          ],
          definedNamesState: null,
        })
      },
      { sessionId: opened.sessionId, sheetId },
    )

    expect(saved.canceled).toBe(false)
    expect(saved.touchedEntries).toContain('xl/tables/table1.xml')
    expect(saved.touchedEntries).toContain('xl/worksheets/sheet1.xml')
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#download-button')).toBeEnabled()

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('web-excel-table-sparkline.xlsx')
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const downloadedZip = await JSZip.loadAsync(await readFile(downloadPath!))
    const tableXml = await downloadedZip.file('xl/tables/table1.xml')?.async('text')
    expect(tableXml).toContain('name="WebTable"')
    expect(tableXml).toContain('displayName="WebTable"')
    expect(tableXml).toContain('ref="A1:B3"')
    expect(tableXml).toContain('<tableColumn id="1" name="Name"/>')
    expect(tableXml).toContain('<tableColumn id="2" name="Value"/>')
    expect(tableXml).toContain('name="TableStyleMedium2"')
    expect(tableXml).toContain('showRowStripes="1"')

    const worksheetXml = await downloadedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(worksheetXml).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )
    expect(worksheetXml).toContain('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
    expect(worksheetXml).toContain(
      '<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">',
    )
    expect(worksheetXml).toContain('<x14:sparklineGroup displayEmptyCellsAs="gap" type="column">')
    expect(worksheetXml).toContain('<x14:colorSeries rgb="FF336699"/>')
    expect(worksheetXml).toContain('<xm:f>Data!B2:B3</xm:f>')
    expect(worksheetXml).toContain('<xm:sqref>D2</xm:sqref>')

    const sheetRels = await downloadedZip
      .file('xl/worksheets/_rels/sheet1.xml.rels')
      ?.async('text')
    expect(sheetRels).toContain('/relationships/table')
    expect(sheetRels).toContain('Target="../tables/table1.xml"')

    const contentTypes = await downloadedZip.file('[Content_Types].xml')?.async('text')
    expect(contentTypes).toContain('PartName="/xl/tables/table1.xml"')
    expect(contentTypes).toContain('spreadsheetml.table+xml')
  })
})
