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
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Clean Save As</t></is></c></row></sheetData>
</worksheet>`,
  }
  for (const [path, content] of Object.entries(parts)) zip.file(path, content)
  return Buffer.from(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })).toString(
    'base64',
  )
}

test.describe('UC Web OS XLSX Host browser shortcuts', () => {
  test.skip(!hostUrl || !sheetsUrl, 'UC Host and Sheets Web preview URLs are required')

  test('Ctrl+Shift+S saves an unchanged workbook as a result file', async ({ page }) => {
    const workbookBase64 = await createWorkbookBase64()

    await page.addInitScript(
      ({ sourceBase64 }) => {
        const decodeBase64 = (input: string): Uint8Array => {
          const binary = atob(input)
          return Uint8Array.from(binary, (character) => character.charCodeAt(0))
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
          saves: [] as Array<{ filename: string; base64: string }>,
          lastWriteMode: 'selected' as 'selected' | 'result',
        }
        ;(window as typeof window & { __ucShortcutMock?: typeof state }).__ucShortcutMock = state

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
                      fileName: 'clean-source.xlsx',
                      nodeId: 'node-source',
                      mode: 'edit',
                      locale: 'zh-CN',
                      file: { nodeId: 'node-source', name: 'clean-source.xlsx' },
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
                    filename: 'clean-source.xlsx',
                    contentType:
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                  }
                  break
                case 'uc.fs.pickSaveDestination':
                  result = { cancelled: false, filename: 'shortcut-copy.xlsx' }
                  break
                case 'uc.fs.saveResultFile': {
                  if (state.lastWriteMode !== 'result') {
                    throw new Error('Shortcut Save As must use result mode.')
                  }
                  const bytes = new Uint8Array(await (params.blob as Blob).arrayBuffer())
                  state.saves.push({ filename: params.filename, base64: encodeBase64(bytes) })
                  result = {
                    nodeId: 'node-copy',
                    filename: params.filename,
                    version: 'copy-v1',
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
                  error: { message: error instanceof Error ? error.message : String(error) },
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
    const openResponse = page.waitForResponse(
      (response) =>
        response.request().method() === 'POST' &&
        response.url().includes('/xlsx-engine/v1/workbooks?name=clean-source.xlsx'),
    )
    await page.goto(
      `${hostUrl}?ucHostOrigin=${encodeURIComponent(hostOrigin)}&sheetsUrl=${encodeURIComponent(
        sheetsUrl!,
      )}&pluginId=thirdparty.plugin.excel-online`,
    )
    await openResponse

    const editorFrame = page.frameLocator('#office-frame')
    const canvas = editorFrame.locator('canvas').first()
    await expect(canvas).toBeVisible({ timeout: 30_000 })
    await canvas.click()
    await page.keyboard.press('Control+Shift+S')

    await expect
      .poll(
        () =>
          page.evaluate(
            () =>
              (window as typeof window & { __ucShortcutMock: any }).__ucShortcutMock.saves.length,
          ),
        { timeout: 30_000 },
      )
      .toBe(1)

    const state = await page.evaluate(() => {
      const value = (window as typeof window & { __ucShortcutMock: any }).__ucShortcutMock
      return { calls: value.calls, saves: value.saves }
    })
    expect(state.calls).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ method: 'uc.fs.pickSaveDestination' }),
        expect.objectContaining({
          method: 'uc.fs.requestSelectedFileAccess',
          writeMode: 'result',
          filename: 'shortcut-copy.xlsx',
        }),
      ]),
    )
    expect(state.saves[0]?.filename).toBe('shortcut-copy.xlsx')

    const zip = await JSZip.loadAsync(Buffer.from(state.saves[0].base64, 'base64'))
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheetXml).toContain('Clean Save As')
  })
})
