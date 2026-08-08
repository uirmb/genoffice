import { createReadStream } from 'node:fs'
import { resolve } from 'node:path'
import { expect, test } from '@playwright/test'
import { parseDocx } from '@genoffice/docx-engine'

const docsWebUrl = process.env.DOCS_WEB_E2E_URL
const hostUrl = process.env.DOCS_WEB_HOST_E2E_URL

test.describe('Docs Web', () => {
  test.skip(!docsWebUrl, 'DOCS_WEB_E2E_URL is only set by the Docs Web browser CI step')

  test('boots the real browser editor and edits a blank document', async ({ page }) => {
    await page.goto(docsWebUrl!)

    const editor = page.locator('.ProseMirror').first()
    await expect(editor).toBeVisible({ timeout: 30_000 })
    await expect(page.locator('text=GenOffice Web failed to start')).toHaveCount(0)

    const text = 'GenOffice Web 浏览器编辑回归'
    await editor.click()
    await page.keyboard.type(text)
    await expect(editor).toContainText(text)

    await expect.poll(() => page.title()).toBeTruthy()
  })

  test('opens, edits, saves, downloads, and reparses a real DOCX through the iframe host', async ({
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

    const editedText = 'Word Web iframe 保存验证中文'
    await editor.click()
    await editor.press('Control+End')
    await page.keyboard.type(editedText)
    await expect(editor).toContainText(editedText)
    await expect(page.locator('#dirty-state')).toHaveText('dirty')

    await page.locator('#save-button').click()
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#dirty-state')).toHaveText('clean')

    const downloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const download = await downloadPromise
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
  })
})
