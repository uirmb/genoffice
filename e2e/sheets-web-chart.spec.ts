import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const sheetsWebUrl = process.env.SHEETS_WEB_E2E_URL
const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL

async function createChartWorkbook(path: string): Promise<void> {
  const zip = new JSZip()
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>B</t></is></c><c r="B3"><v>20</v></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
    'xl/drawings/drawing1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`,
    'xl/charts/chart1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original Chart</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/><c:tx><c:v>Value</c:v></c:tx>
          <c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="150"/><c:overlap val="-27"/><c:axId val="111"/><c:axId val="222"/>
      </c:barChart>
      <c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>
      <c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

test.describe('Sheets Web chart edits', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('edits a linked chart part and preserves its drawing relationship', async ({ page }) => {
    test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is required for the iframe host flow')

    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-web-chart-'))
    const workbookPath = join(directory, 'web-excel-chart.xlsx')
    await createChartWorkbook(workbookPath)

    await page.goto(hostUrl!)
    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })

    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=web-excel-chart.xlsx'),
    )
    await page.locator('#xlsx-picker').setInputFiles(workbookPath)
    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      sheets: Array<{ id: string; name: string }>
      visuals?: Array<{ kind: string; chartPath?: string }>
    }
    expect(opened.sheets[0]?.name).toBe('Data')
    expect(opened.visuals?.some((visual) => visual.kind === 'chart')).toBe(true)

    const saved = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          sessionId: payload.sessionId,
          mode: 'save',
          edits: [],
          structuralOps: [],
          chartEdits: [{ chartPath: 'xl/charts/chart1.xml', title: 'Web Chart Saved' }],
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
      { sessionId: opened.sessionId },
    )

    expect(saved.canceled).toBe(false)
    expect(saved.touchedEntries).toContain('xl/charts/chart1.xml')
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#download-button')).toBeEnabled()

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('web-excel-chart.xlsx')
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const downloadedZip = await JSZip.loadAsync(await readFile(downloadPath!))
    const chartXml = await downloadedZip.file('xl/charts/chart1.xml')?.async('text')
    expect(chartXml).toContain('<a:t>Web Chart Saved</a:t>')
    expect(chartXml).not.toContain('Original Chart')
    expect(chartXml).toContain('<c:f>Data!$A$2:$A$3</c:f>')
    expect(chartXml).toContain('<c:f>Data!$B$2:$B$3</c:f>')

    const worksheetXml = await downloadedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(worksheetXml).toContain('<drawing r:id="rId1"/>')
    const drawingRels = await downloadedZip.file('xl/drawings/_rels/drawing1.xml.rels')?.async('text')
    expect(drawingRels).toContain('Target="../charts/chart1.xml"')
    const drawingXml = await downloadedZip.file('xl/drawings/drawing1.xml')?.async('text')
    expect(drawingXml).toContain('<c:chart r:id="rId1"/>')
  })
})
