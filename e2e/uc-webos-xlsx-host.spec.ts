import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

const hostUrl = process.env.UC_WEBOS_XLSX_HOST_E2E_URL
const sheetsUrl = process.env.SHEETS_WEB_E2E_URL

async function createWorkbookBase64(): Promise<string> {
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
  <dimension ref="A1:B1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>UC Source</t></is></c><c r="B1"><v>42</v></c></row></sheetData>
</worksheet>`,
  }
  for (const [path, content] of Object.entries(parts)) zip.file(path, content)
  return Buffer.from(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })).toString(
    'base64',
  )
}

const emptySaveRequest = {
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

test.describe('UC Web OS XLSX Host', () => {
  test.skip(!hostUrl || !sheetsUrl, 'UC Host and Sheets Web preview URLs are required')

  test('opens through UC selected-file RPC and saves selected/result files', async ({ page }) => {
    const workbookBase64 = await createWorkbookBase64()

    await page.addInitScript(
      ({ sourceBase64 }) => {
        const decodeBase64 = (input: string): Uint8Array => {
          const binary = atob(input)
          const bytes = new Uint8Array(binary.length)
          for (let index = 0; index < binary.length; index += 1) {
            bytes[index] = binary.charCodeAt(index)
          }
          return bytes
        }
        const encodeBase64 = (bytes: Uint8Array): string => {
          let binary = ''
          for (let offset = 0; offset < bytes.length; offset += 0x8000) {
            binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
          }
          return btoa(binary)
        }

        const state = {
          calls: [] as Array<{ method: string; writeMode?: string; filename?: string }>,
          saves: [] as Array<{
            filename: string
            writeMode: string
            nodeId: string
            base64: string
          }>,
          lastWriteMode: 'selected' as 'selected' | 'result',
        }
        ;(window as typeof window & { __ucMock?: typeof state }).__ucMock = state

        window.addEventListener('message', (event) => {
          const message = event.data as any
          if (event.source !== window || message?.type !== 'uc-plugin-rpc-request') return

          void (async () => {
            try {
              const params = message.params || {}
              state.calls.push({
                method: message.method,
                ...(params.writeMode ? { writeMode: params.writeMode } : {}),
                ...(params.filename ? { filename: params.filename } : {}),
              })

              let result: unknown
              switch (message.method) {
                case 'uc.ready':
                  result = { ok: true }
                  break
                case 'uc.host.getLaunchParams':
                  result = {
                    launchParams: {
                      fileName: 'uc-source.xlsx',
                      nodeId: 'node-source',
                      mode: 'edit',
                      locale: 'zh-CN',
                      file: { nodeId: 'node-source', name: 'uc-source.xlsx' },
                    },
                  }
                  break
                case 'uc.fs.requestSelectedFileAccess':
                  state.lastWriteMode = params.writeMode
                  result =
                    params.writeMode === 'result'
                      ? {
                          nodeId: 'node-source',
                          resultNodeId: 'node-copy',
                          filename: params.filename,
                          version: 'copy-v1',
                        }
                      : {
                          nodeId: 'node-source',
                          filename: params.filename,
                          version: 'source-v1',
                        }
                  break
                case 'uc.fs.readSelectedFile':
                  result = {
                    blob: new Blob([decodeBase64(sourceBase64)], {
                      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    }),
                    filename: 'uc-source.xlsx',
                    contentType:
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                  }
                  break
                case 'uc.fs.pickSaveDestination':
                  result = {
                    cancelled: false,
                    filename: 'uc-copy.xlsx',
                    folderName: 'Documents',
                  }
                  break
                case 'uc.fs.saveResultFile': {
                  const blob = params.blob as Blob
                  const bytes = new Uint8Array(await blob.arrayBuffer())
                  const nodeId = state.lastWriteMode === 'result' ? 'node-copy' : 'node-source'
                  state.saves.push({
                    filename: params.filename,
                    writeMode: state.lastWriteMode,
                    nodeId,
                    base64: encodeBase64(bytes),
                  })
                  result = {
                    nodeId,
                    filename: params.filename,
                    version: state.lastWriteMode === 'result' ? 'copy-v1' : 'source-v2',
                  }
                  break
                }
                default:
                  throw new Error(`Unexpected UC RPC: ${message.method}`)
              }

              window.postMessage(
                {
                  type: 'uc-plugin-rpc-response',
                  id: message.id,
                  pluginId: message.pluginId,
                  result,
                },
                window.location.origin,
              )
            } catch (error) {
              window.postMessage(
                {
                  type: 'uc-plugin-rpc-response',
                  id: message.id,
                  pluginId: message.pluginId,
                  error: {
                    message: error instanceof Error ? error.message : String(error),
                  },
                },
                window.location.origin,
              )
            }
          })()
        })
      },
      { sourceBase64: workbookBase64 },
    )

    const hostOrigin = new URL(hostUrl!).origin
    const openResponsePromise = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=uc-source.xlsx'),
    )
    await page.goto(
      `${hostUrl}?ucHostOrigin=${encodeURIComponent(hostOrigin)}&sheetsUrl=${encodeURIComponent(
        sheetsUrl!,
      )}&pluginId=thirdparty.plugin.excel-online`,
    )

    const opened = (await (await openResponsePromise).json()) as {
      sessionId: string
      name: string
      sheets: Array<{ id: string; name: string }>
    }
    expect(opened.name).toBe('uc-source.xlsx')
    expect(opened.sheets[0]?.name).toBe('Data')

    const editorFrame = page.frameLocator('#office-frame')
    await expect(editorFrame.locator('canvas').first()).toBeVisible({ timeout: 30_000 })
    const sheetId = opened.sheets[0]!.id

    const firstSave = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          ...payload.empty,
          sessionId: payload.sessionId,
          mode: 'save',
          edits: [
            {
              sheetId: payload.sheetId,
              row: 0,
              column: 0,
              writeValue: true,
              value: 'UC Saved',
            },
          ],
        })
      },
      {
        sessionId: opened.sessionId,
        sheetId,
        empty: emptySaveRequest,
      },
    )
    expect(firstSave.canceled).toBe(false)

    const firstMock = await page.evaluate(() => {
      const state = (window as typeof window & { __ucMock: any }).__ucMock
      return { calls: state.calls, saves: state.saves }
    })
    expect(firstMock.calls.map((call: any) => call.method)).toEqual(
      expect.arrayContaining([
        'uc.ready',
        'uc.host.getLaunchParams',
        'uc.fs.requestSelectedFileAccess',
        'uc.fs.readSelectedFile',
        'uc.fs.saveResultFile',
      ]),
    )
    expect(firstMock.saves).toHaveLength(1)
    expect(firstMock.saves[0]).toMatchObject({
      filename: 'uc-source.xlsx',
      writeMode: 'selected',
      nodeId: 'node-source',
    })

    const firstSavedZip = await JSZip.loadAsync(Buffer.from(firstMock.saves[0].base64, 'base64'))
    const firstSheetXml = await firstSavedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(firstSheetXml).toContain('UC Saved')

    const saveAs = await editorFrame.locator('body').evaluate(
      async (_body, payload) => {
        const api = (window as typeof window & { desktopApi: any }).desktopApi
        return api.saveWorkbookEdits({
          ...payload.empty,
          sessionId: payload.sessionId,
          mode: 'save-as',
          edits: [],
        })
      },
      { sessionId: firstSave.file.sessionId, empty: emptySaveRequest },
    )
    expect(saveAs.canceled).toBe(false)

    const finalMock = await page.evaluate(() => {
      const state = (window as typeof window & { __ucMock: any }).__ucMock
      return { calls: state.calls, saves: state.saves }
    })
    expect(finalMock.saves).toHaveLength(2)
    expect(finalMock.saves[1]).toMatchObject({
      filename: 'uc-copy.xlsx',
      writeMode: 'result',
      nodeId: 'node-copy',
    })
    expect(finalMock.calls).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ method: 'uc.fs.pickSaveDestination' }),
        expect.objectContaining({
          method: 'uc.fs.requestSelectedFileAccess',
          writeMode: 'result',
          filename: 'uc-copy.xlsx',
        }),
      ]),
    )

    const copiedZip = await JSZip.loadAsync(Buffer.from(finalMock.saves[1].base64, 'base64'))
    const copiedSheetXml = await copiedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(copiedSheetXml).toContain('UC Saved')
  })
})
