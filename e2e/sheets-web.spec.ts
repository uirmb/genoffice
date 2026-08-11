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
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="RemoveMe" sheetId="2" r:id="rId2"/></sheets>
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
  <dimension ref="A1:C1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Browser Original</t></is></c><c r="B1"><v>42</v></c><c r="C1"><f>B1*2</f><v>84</v></c></row></sheetData>
</worksheet>`,
    'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Remove this sheet</t></is></c></row></sheetData>
</worksheet>`,
  }

  for (const [name, content] of Object.entries(parts)) zip.file(name, content)
  const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  await writeFile(path, bytes)
}

test.describe('Sheets Web', () => {
  test.skip(!sheetsWebUrl, 'SHEETS_WEB_E2E_URL is only set by the Sheets Web browser CI step')

  test('preserves formulas, worksheet journals, and sheet management through the iframe host', async ({
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
      sheets: Array<{ id: string; name: string; hidden?: boolean }>
    }
    expect(opened.sessionId).toBeTruthy()
    expect(opened.sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'RemoveMe'])

    const sheetId = opened.sheets[0]!.id
    const removeSheetId = opened.sheets[1]!.id
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
      '=B1*2',
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

    const addedSheetId = 'web-added-sheet'
    const copySheetId = 'web-copy-sheet'
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
            {
              sheetId: payload.sheetId,
              row: 0,
              column: 1,
              writeValue: true,
              value: 50,
            },
            {
              sheetId: payload.addedSheetId,
              row: 0,
              column: 0,
              writeValue: true,
              value: 'Added Web Sheet',
            },
            {
              sheetId: payload.copySheetId,
              row: 1,
              column: 0,
              writeValue: true,
              value: 'Edited Copy',
            },
          ],
          structuralOps: [],
          chartEdits: [],
          visualEdits: [],
          visualAdditions: [],
          tableAdditions: [],
          pivotAdditions: [],
          sheetOps: [
            { kind: 'rename-sheet', sheetId: payload.sheetId, newName: 'Renamed' },
            { kind: 'add-sheet', sheetId: payload.addedSheetId, name: 'Added' },
            {
              kind: 'duplicate-sheet',
              sheetId: payload.copySheetId,
              name: 'Copy',
              sourceSheetId: payload.sheetId,
            },
            { kind: 'remove-sheet', sheetId: payload.removeSheetId },
            { kind: 'set-sheet-hidden', sheetId: payload.addedSheetId, hidden: true },
            { kind: 'reorder-sheets' },
          ],
          sheetOrder: [payload.copySheetId, payload.sheetId, payload.addedSheetId],
          filterStates: [
            {
              sheetId: payload.sheetId,
              filter: {
                range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
                columns: [{ colId: 0, values: ['Browser Saved'] }],
              },
              hiddenRows: [],
              visibilityRange: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
            },
          ],
          hyperlinkEdits: [
            {
              sheetId: payload.sheetId,
              row: 0,
              column: 0,
              target: '#Sheet1!B1',
            },
          ],
          cfStates: [
            {
              sheetId: payload.sheetId,
              rules: [
                {
                  ranges: [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }],
                  stopIfTrue: false,
                  rule: {
                    type: 'highlightCell',
                    subType: 'formula',
                    value: '=B1>0',
                    style: {},
                  },
                },
              ],
            },
          ],
          dvStates: [
            {
              sheetId: payload.sheetId,
              rules: [
                {
                  ranges: [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }],
                  rule: {
                    type: 'whole',
                    operator: 'between',
                    formula1: '0',
                    formula2: '100',
                    allowBlank: true,
                  },
                },
              ],
            },
          ],
          pageSetupStates: [
            {
              sheetId: payload.sheetId,
              orientation: 'landscape',
              printGridlines: true,
            },
          ],
          noteStates: [
            {
              sheetId: payload.sheetId,
              notes: [{ row: 0, column: 0, author: 'GenOffice', text: 'Web note' }],
            },
          ],
          formulaValues: [
            {
              sheetId: payload.sheetId,
              row: 0,
              column: 2,
              value: 100,
            },
          ],
          pivotCacheRefreshPaths: [],
          pivotRefreshUpdates: [],
          sheetProtections: [{ sheetId: payload.sheetId, protected: true }],
          sparklineAdditions: [],
          definedNamesState: null,
        })
      },
      {
        sessionId: opened.sessionId,
        sheetId,
        removeSheetId,
        addedSheetId,
        copySheetId,
      },
    )

    expect(saved.canceled).toBe(false)
    expect(saved.touchedEntries).toContain('xl/worksheets/sheet1.xml')
    expect(saved.file.sheets.map((sheet: any) => sheet.name)).toEqual(['Copy', 'Renamed', 'Added'])
    expect(saved.file.sheets.find((sheet: any) => sheet.name === 'Added')?.hidden).toBe(true)
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#file-name')).toHaveText('web-excel-browser-e2e.xlsx')
    await expect(page.locator('#download-button')).toBeEnabled()

    const renamedSheetId = saved.file.sheets.find((sheet: any) => sheet.name === 'Renamed')!.id
    const copySavedSheetId = saved.file.sheets.find((sheet: any) => sheet.name === 'Copy')!.id
    const addedSavedSheetId = saved.file.sheets.find((sheet: any) => sheet.name === 'Added')!.id

    const savedRange = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readWorkbookRange({
          sessionId: payload.sessionId,
          sheetId: payload.sheetId,
          range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
        })
      },
      { sessionId: saved.file.sessionId, sheetId: renamedSheetId },
    )
    expect(savedRange.cells.find((cell: any) => cell.row === 0 && cell.column === 0)?.value).toBe(
      'Browser Saved',
    )
    expect(savedRange.cells.find((cell: any) => cell.row === 0 && cell.column === 1)?.value).toBe(50)
    const savedFormula = savedRange.cells.find((cell: any) => cell.row === 0 && cell.column === 2)
    expect(savedFormula?.formula).toBe('=B1*2')
    expect(savedFormula?.value).toBe(100)

    const copyRange = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readWorkbookRange({
          sessionId: payload.sessionId,
          sheetId: payload.sheetId,
          range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
        })
      },
      { sessionId: saved.file.sessionId, sheetId: copySavedSheetId },
    )
    expect(copyRange.cells.find((cell: any) => cell.row === 0 && cell.column === 0)?.value).toBe(
      'Browser Original',
    )
    expect(copyRange.cells.find((cell: any) => cell.row === 1 && cell.column === 0)?.value).toBe(
      'Edited Copy',
    )

    const addedRange = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.readWorkbookRange({
          sessionId: payload.sessionId,
          sheetId: payload.sheetId,
          range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        })
      },
      { sessionId: saved.file.sessionId, sheetId: addedSavedSheetId },
    )
    expect(addedRange.cells.find((cell: any) => cell.row === 0 && cell.column === 0)?.value).toBe(
      'Added Web Sheet',
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
    expect(worksheetXml).toMatch(/<c r="C1"[^>]*><f>B1\*2<\/f><v>100<\/v><\/c>/)
    expect(worksheetXml).toContain('<sheetProtection sheet="1" objects="1" scenarios="1"/>')
    expect(worksheetXml).toContain('<autoFilter ref="A1:C1">')
    expect(worksheetXml).toContain('<hyperlink ref="A1" location="Renamed!B1"/>')
    expect(worksheetXml).toContain('<conditionalFormatting sqref="B1">')
    expect(worksheetXml).toContain('<formula>B1&gt;0</formula>')
    expect(worksheetXml).toContain('<dataValidations count="1">')
    expect(worksheetXml).toMatch(
      /<dataValidation\b[^>]*\btype="whole"[^>]*\ballowBlank="1"[^>]*\bsqref="B1"/,
    )
    expect(worksheetXml).toContain('<formula1>0</formula1>')
    expect(worksheetXml).toContain('<formula2>100</formula2>')
    expect(worksheetXml).toMatch(/<pageSetup\b[^>]*\borientation="landscape"/)
    expect(worksheetXml).toMatch(/<printOptions\b[^>]*\bgridLines="1"/)
    expect(worksheetXml).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )
    expect(worksheetXml).toContain('<legacyDrawing r:id=')

    const commentsXml = await downloadedZip.file('xl/comments1.xml')?.async('text')
    expect(commentsXml).toContain('<author>GenOffice</author>')
    expect(commentsXml).toContain('<comment ref="A1" authorId="0">')
    expect(commentsXml).toContain('Web note')
    expect(downloadedZip.file('xl/drawings/vmlDrawing1.vml')).not.toBeNull()

    const worksheetRels = await downloadedZip
      .file('xl/worksheets/_rels/sheet1.xml.rels')
      ?.async('text')
    expect(worksheetRels).toContain('/relationships/comments')
    expect(worksheetRels).toContain('/relationships/vmlDrawing')

    expect(downloadedZip.file('xl/worksheets/sheet2.xml')).toBeNull()
    const workbookXml = await downloadedZip.file('xl/workbook.xml')?.async('text')
    expect(workbookXml).toMatch(
      /<sheets>\s*<sheet\b[^>]*name="Copy"[\s\S]*?<sheet\b[^>]*name="Renamed"[\s\S]*?<sheet\b[^>]*name="Added"[^>]*state="hidden"[^>]*\/>\s*<\/sheets>/,
    )
    expect(workbookXml).not.toContain('name="RemoveMe"')

    const workbookRels = await downloadedZip.file('xl/_rels/workbook.xml.rels')?.async('text')
    expect(workbookRels).not.toContain('Target="worksheets/sheet2.xml"')

    const contentTypes = await downloadedZip.file('[Content_Types].xml')?.async('text')
    expect(contentTypes).toContain('spreadsheetml.comments+xml')
    expect(contentTypes).toContain('Extension="vml"')
    expect(contentTypes).not.toContain('PartName="/xl/worksheets/sheet2.xml"')
  })
})
