import { expect, test } from '@playwright/test'

const docsWebUrl = process.env.DOCS_WEB_E2E_URL

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
})
