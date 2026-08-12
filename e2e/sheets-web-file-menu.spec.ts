import { resolve } from 'node:path'

import { expect, test, type FrameLocator, type Page } from '@playwright/test'

const hostUrl = process.env.SHEETS_WEB_HOST_E2E_URL
const fixture = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

async function openHost(page: Page): Promise<FrameLocator> {
  if (!hostUrl) throw new Error('SHEETS_WEB_HOST_E2E_URL is required.')
  await page.goto(hostUrl)
  const editor = page.frameLocator('#office-frame')
  await expect(editor.locator('canvas').first()).toBeVisible({ timeout: 30_000 })
  await expect(editor.locator('text=GenOffice Sheets Web failed to start')).toHaveCount(0)
  await expect(editor.locator('.sheets-web-file-menu-root .ribbon-tab-file')).toHaveText('文件', {
    timeout: 15_000,
  })
  return editor
}

async function openFixture(page: Page): Promise<void> {
  await page.locator('#xlsx-picker').setInputFiles(fixture)
  await expect(page.locator('#file-name')).toContainText('compatibility-basic.xlsx', {
    timeout: 15_000,
  })
}

async function openFileMenu(editor: FrameLocator): Promise<void> {
  const trigger = editor.locator('.sheets-web-file-menu-root .ribbon-tab-file')
  await trigger.click()
  await expect(editor.locator('.sheets-web-file-menu-root .file-menu')).toBeVisible()
}

async function editFirstCell(page: Page, editor: FrameLocator): Promise<void> {
  const canvases = editor.locator('canvas')
  const count = await canvases.count()
  for (let index = 0; index < count; index += 1) {
    const box = await canvases.nth(index).boundingBox()
    if (!box || box.width < 500 || box.height < 300) continue
    // The worksheet canvas includes a narrow row/column header. A1 is safely
    // inside this offset on the Univer grid used by Sheets Web.
    await page.mouse.click(box.x + 90, box.y + 38)
    await page.keyboard.type('File Menu Dirty', { delay: 20 })
    await page.keyboard.press('Enter')
    await expect(editor.locator('.ribbon-tabs .qa-btn').first()).toBeEnabled({ timeout: 10_000 })
    return
  }
  throw new Error('Worksheet canvas was not found.')
}

async function installCloseMessageRecorder(page: Page): Promise<void> {
  await page.evaluate(() => {
    const target = window as Window & {
      __sheetsCloseMessages?: Array<{
        type: string
        requestId?: string
        reason?: string
      }>
    }
    target.__sheetsCloseMessages = []
    window.addEventListener('message', (event) => {
      const message = event.data as {
        type?: string
        requestId?: string
        payload?: { reason?: string }
      }
      if (
        message?.type !== 'office:close-request' &&
        message?.type !== 'office:close-cancelled'
      ) {
        return
      }
      target.__sheetsCloseMessages?.push({
        type: message.type,
        ...(message.requestId ? { requestId: message.requestId } : {}),
        ...(message.payload?.reason ? { reason: message.payload.reason } : {}),
      })
    })
  })
}

async function sendWindowCloseRequest(page: Page, requestId: string): Promise<void> {
  await page.evaluate((id) => {
    const frame = document.querySelector<HTMLIFrameElement>('#office-frame')
    if (!frame?.contentWindow) throw new Error('Sheets iframe is not available.')
    const targetOrigin = new URL(frame.src, window.location.href).origin
    frame.contentWindow.postMessage(
      {
        protocol: 1,
        type: 'office:request-close',
        requestId: id,
        payload: { reason: 'window-close' },
      },
      targetOrigin,
    )
  }, requestId)
}

async function expectCloseMessage(
  page: Page,
  type: 'office:close-request' | 'office:close-cancelled',
  requestId: string,
  reason: string,
): Promise<void> {
  await expect
    .poll(
      () =>
        page.evaluate(
          ({ expectedType, expectedRequestId, expectedReason }) => {
            const messages = (
              window as Window & {
                __sheetsCloseMessages?: Array<{
                  type: string
                  requestId?: string
                  reason?: string
                }>
              }
            ).__sheetsCloseMessages
            return (
              messages?.some(
                (message) =>
                  message.type === expectedType &&
                  message.requestId === expectedRequestId &&
                  message.reason === expectedReason,
              ) ?? false
            )
          },
          {
            expectedType: type,
            expectedRequestId: requestId,
            expectedReason: reason,
          },
        ),
      { timeout: 20_000 },
    )
    .toBe(true)
}

test.describe('Sheets Web deployable File menu', () => {
  test.skip(!hostUrl, 'SHEETS_WEB_HOST_E2E_URL is only set by the Sheets Web File UI gate')

  test('matches the Word/PPT File position and hides Web AI product chrome', async ({ page }) => {
    const editor = await openHost(page)
    const fileButton = editor.locator('.sheets-web-file-menu-root .ribbon-tab-file')
    const saveButton = editor.locator('.ribbon-tabs .qa-btn').first()

    const fileBox = await fileButton.boundingBox()
    const saveBox = await saveButton.boundingBox()
    expect(fileBox).not.toBeNull()
    expect(saveBox).not.toBeNull()
    expect(fileBox!.x).toBeLessThan(saveBox!.x)
    expect(Math.abs(fileBox!.y - saveBox!.y)).toBeLessThan(8)

    await expect(editor.locator('.ai-entry:visible')).toHaveCount(0)
    await expect(editor.locator('.copilot:visible')).toHaveCount(0)
    await expect(editor.locator('.autosave-toggle:visible')).toHaveCount(0)
    await expect(editor.locator('.workbook-status')).not.toContainText(/AI|Genspark/i)

    await openFileMenu(editor)
    const visibleItems = editor.locator('.sheets-web-file-menu-root .file-menu button:visible')
    await expect(visibleItems).toHaveCount(6)
    await expect(visibleItems.nth(0).locator('span').first()).toHaveText('打开')
    await expect(visibleItems.nth(0).locator('.file-menu-key')).toHaveText('Ctrl+O')
    await expect(visibleItems.nth(1).locator('span').first()).toHaveText('保存')
    await expect(visibleItems.nth(1).locator('.file-menu-key')).toHaveText('Ctrl+S')
    await expect(visibleItems.nth(2).locator('span').first()).toHaveText('另存为')
    await expect(visibleItems.nth(2).locator('.file-menu-key')).toHaveText('Ctrl+Shift+S')
    await expect(visibleItems.nth(3)).toHaveText('保存历史版本')
    await expect(visibleItems.nth(4)).toHaveText('导出为 XLSX')
    await expect(visibleItems.nth(5)).toHaveText('退出')
  })

  test('runs history, XLSX export, and unchanged Save As through the Host', async ({ page }) => {
    const editor = await openHost(page)
    await openFixture(page)

    await openFileMenu(editor)
    await editor.locator('.file-menu-save-history').click()
    await expect(page.locator('#host-state')).toContainText('history saved (1)', { timeout: 20_000 })
    await expect(page.locator('#file-name')).toContainText('compatibility-basic.xlsx')

    const downloadPromise = page.waitForEvent('download')
    await openFileMenu(editor)
    await editor.locator('.file-menu-export-xlsx').click()
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('compatibility-basic.xlsx')
    await expect(page.locator('#file-name')).toContainText('compatibility-basic.xlsx')

    page.once('dialog', async (dialog) => {
      expect(dialog.type()).toBe('prompt')
      await dialog.accept('file-menu-copy.xlsx')
    })
    await openFileMenu(editor)
    await editor.locator('.file-menu-save-as').click()
    await expect(page.locator('#file-name')).toHaveText('file-menu-copy.xlsx', { timeout: 20_000 })
    await expect(page.locator('#host-state')).toContainText('saved')
  })

  test('uses Save / Discard / Cancel for dirty File Exit', async ({ page }) => {
    const editor = await openHost(page)
    await openFixture(page)
    await editFirstCell(page, editor)

    await openFileMenu(editor)
    await editor.locator('.file-menu-exit').click()
    const dialog = editor.locator('.file-exit-dialog')
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('button')).toHaveText([
      '取消',
      '放弃更改并退出',
      '保存并退出',
    ])

    await dialog.getByRole('button', { name: '取消' }).click()
    await expect(dialog).toHaveCount(0)

    await openFileMenu(editor)
    await editor.locator('.file-menu-exit').click()
    await editor.locator('.file-exit-dialog').getByRole('button', { name: '放弃更改并退出' }).click()
    await expect(page.locator('#host-state')).toContainText('close requested', { timeout: 10_000 })
  })

  test('saves dirty workbook before granting File Exit', async ({ page }) => {
    const editor = await openHost(page)
    await openFixture(page)
    await editFirstCell(page, editor)

    await openFileMenu(editor)
    await editor.locator('.file-menu-exit').click()
    const dialog = editor.locator('.file-exit-dialog')
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: '保存并退出' }).click()

    await expect(editor.locator('.ribbon-tabs .qa-btn').first()).toBeDisabled({ timeout: 20_000 })
    await expect(page.locator('#dirty-state')).toHaveText('clean', { timeout: 20_000 })
    await expect(page.locator('#host-state')).toContainText('close requested', { timeout: 20_000 })
  })

  test('returns the same requestId for a clean Host window close', async ({ page }) => {
    await openHost(page)
    await installCloseMessageRecorder(page)

    await sendWindowCloseRequest(page, 'close-clean-1')
    await expectCloseMessage(page, 'office:close-request', 'close-clean-1', 'window-close')
    await expect(page.locator('#host-state')).toContainText('close requested', { timeout: 10_000 })
  })

  test('cancels one dirty Host close transaction and ignores duplicate requests', async ({ page }) => {
    const editor = await openHost(page)
    await openFixture(page)
    await editFirstCell(page, editor)
    await installCloseMessageRecorder(page)

    await sendWindowCloseRequest(page, 'close-dirty-1')
    await sendWindowCloseRequest(page, 'close-dirty-duplicate')

    const dialog = editor.locator('.file-exit-dialog')
    await expect(dialog).toBeVisible()
    await expect(editor.locator('.file-exit-dialog')).toHaveCount(1)
    await dialog.getByRole('button', { name: '取消' }).click()

    await expectCloseMessage(page, 'office:close-cancelled', 'close-dirty-1', 'user-cancelled')
    await page.waitForTimeout(250)
    expect(
      await page.evaluate(() =>
        (
          window as Window & {
            __sheetsCloseMessages?: Array<{ requestId?: string }>
          }
        ).__sheetsCloseMessages?.some(
          (message) => message.requestId === 'close-dirty-duplicate',
        ) ?? false,
      ),
    ).toBe(false)

    await sendWindowCloseRequest(page, 'close-dirty-2')
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: '放弃更改并退出' }).click()
    await expectCloseMessage(page, 'office:close-request', 'close-dirty-2', 'window-close')
  })

  test('saves dirty workbook before granting a correlated Host window close', async ({ page }) => {
    const editor = await openHost(page)
    await openFixture(page)
    await editFirstCell(page, editor)
    await installCloseMessageRecorder(page)

    await sendWindowCloseRequest(page, 'close-save-1')
    const dialog = editor.locator('.file-exit-dialog')
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: '保存并退出' }).click()

    await expect(page.locator('#dirty-state')).toHaveText('clean', { timeout: 20_000 })
    await expectCloseMessage(page, 'office:close-request', 'close-save-1', 'window-close')
  })
})
