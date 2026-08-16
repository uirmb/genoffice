import { createReadStream } from 'node:fs'
import { resolve } from 'node:path'
import { expect, test } from '@playwright/test'
import { parseDocx } from '@genoffice/docx-engine'

const docsWebUrl = process.env.DOCS_WEB_E2E_URL
const hostUrl = process.env.DOCS_WEB_HOST_E2E_URL

test.describe('Docs Web', () => {
  test.skip(!docsWebUrl, 'DOCS_WEB_E2E_URL is only set by the Docs Web browser CI step')

  test('boots the productized browser editor and edits a blank document', async ({ page }) => {
    const unexpectedDownloads: string[] = []
    page.on('download', (download) => unexpectedDownloads.push(download.suggestedFilename()))

    await page.goto(docsWebUrl!)

    const editor = page.locator('.ProseMirror').first()
    await expect(editor).toBeVisible({ timeout: 30_000 })
    await expect(page.locator('text=GenOffice Web failed to start')).toHaveCount(0)
    await expect(page.locator('.autosave-toggle')).toBeHidden()
    await expect(page.locator('.ribbon-group:has(.ai-entry)')).toBeHidden()
    await expect(page.locator('.ai-dock')).toBeHidden()
    await expect(page.locator('html')).toHaveClass(/office-page-crop-marks/)

    const text = 'GenOffice Web 浏览器编辑回归'
    await editor.click()
    await page.keyboard.type(text)
    await expect(editor).toContainText(text)

    // A pathless Web document must never turn background recovery or the
    // legacy AI-complete silent-save hook into a browser download.
    await page.evaluate(() => window.dispatchEvent(new Event('ai-docs-run-done')))
    await page.waitForTimeout(31_000)
    expect(unexpectedDownloads).toEqual([])

    await expect.poll(() => page.title()).toBeTruthy()
  })

  test('opens, edits, saves, saves as, downloads, and reparses a real DOCX through the iframe host', async ({
    page,
  }) => {
    test.skip(!hostUrl, 'DOCS_WEB_HOST_E2E_URL is required for the iframe host flow')

    await page.goto(hostUrl!)
    await page.locator('#docx-picker').setInputFiles(resolve('fixtures/generated/simple.docx'))

    const editorFrame = page.frameLocator('#office-frame')
    const editor = editorFrame.locator('.ProseMirror').first()
    await expect(editor).toBeVisible({ timeout: 30_000 })
    await expect(editor).toContainText('标题')
    await expect(editor).toContainText('第一段。')
    await expect(editorFrame.locator('.autosave-toggle')).toBeHidden()
    await expect(editorFrame.locator('.ribbon-group:has(.ai-entry)')).toBeHidden()
    await expect(editorFrame.locator('.ai-dock')).toBeHidden()

    // The embedded Web product owns a complete host-facing File menu. The
    // platform decides how Open/history/export/close are ultimately handled.
    await editorFrame.locator('.ribbon-tab-file').click()
    await expect(editorFrame.locator('.file-menu-open')).toBeVisible()
    await expect(editorFrame.locator('.file-menu-save')).toBeVisible()
    await expect(editorFrame.locator('.file-menu-save')).toBeDisabled()
    await expect(editorFrame.locator('.file-menu-save-as')).toBeVisible()
    await expect(editorFrame.locator('.file-menu-save-history')).toHaveText('存为新的历史版本')
    await expect(editorFrame.locator('.file-menu-export-docx')).toHaveText('下载到本地')
    await expect(editorFrame.locator('.file-menu-exit')).toBeVisible()
    await editorFrame.locator('.ribbon-tab-file').click()

    const editedText = 'Word Web iframe 保存验证中文'
    await editor.click()
    await editor.press('Control+End')
    await page.keyboard.type(editedText)
    await expect(editor).toContainText(editedText)
    await expect(page.locator('#dirty-state')).toHaveText('dirty')
    await editorFrame.locator('.ribbon-tab-file').click()
    await expect(editorFrame.locator('.file-menu-save')).toBeEnabled()
    await editorFrame.locator('.ribbon-tab-file').click()

    await page.locator('#save-button').click()
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#dirty-state')).toHaveText('clean')
    await editorFrame.locator('.ribbon-tab-file').click()
    await expect(editorFrame.locator('.file-menu-save')).toBeDisabled()
    await editorFrame.locator('.ribbon-tab-file').click()

    const saveAsText = '另存为链路验证'
    await editor.click()
    await editor.press('Control+End')
    await page.keyboard.type(saveAsText)
    await expect(page.locator('#dirty-state')).toHaveText('dirty')

    page.once('dialog', async (dialog) => {
      expect(dialog.type()).toBe('prompt')
      await dialog.accept('Word-Web-另存为验证.docx')
    })
    await editorFrame.locator('.ribbon-tab-file').click()
    await editorFrame.locator('.file-menu-save-as').click()
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#file-name')).toHaveText('Word-Web-另存为验证.docx')
    await expect(page.locator('#dirty-state')).toHaveText('clean')

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
    expect(download.suggestedFilename()).toBe('Word-Web-另存为验证.docx')
    const downloadPath = await download.path()
    expect(downloadPath).toBeTruthy()

    const chunks: Buffer[] = []
    for await (const chunk of createReadStream(downloadPath!)) chunks.push(Buffer.from(chunk))
    const saved = new Uint8Array(Buffer.concat(chunks))
    const reparsed = await parseDocx(saved)
    const text = reparsed.blocks
      .filter((block) => !block.hidden)
      .flatMap((block) => block.runs ?? [])
      .map((run) => run.text)
      .join('')

    expect(text).toContain('标题')
    expect(text).toContain(editedText)
    expect(text).toContain(saveAsText)
  })
})
