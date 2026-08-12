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
    expect(
      (await visibleItems.allTextContents()).map((value) => value.replace(/\s+/g, ' ').trim()),
    ).toEqual([
      '打开 Ctrl+O',
      '保存 Ctrl+S',
      '另存为 Ctrl+Shift+S',
      '保存历史版本',
      '导出为 XLSX',
      '退出',
    ])
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
})
