import { expect, test } from '@playwright/test'

const MARKDOWN_WEB_URL = process.env.MARKDOWN_WEB_E2E_URL ?? 'http://127.0.0.1:5277/'
const MARKDOWN_WEB_HOST_URL = process.env.MARKDOWN_WEB_HOST_E2E_URL ?? 'http://127.0.0.1:8084/'

const ONE_PIXEL_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2n0AAAAAASUVORK5CYII=',
  'base64',
)

test('standalone Markdown Web starts as an editable AI-free document', async ({ page }) => {
  await page.goto(MARKDOWN_WEB_URL)

  await expect(page.locator('.doc-editor')).toBeVisible()
  await expect(page.locator('.doc-editor')).toHaveAttribute('contenteditable', 'true')
  await expect(page.locator('.ai-entry:visible')).toHaveCount(0)
  await expect(page.locator('.ai-dock:visible')).toHaveCount(0)
})

test('embedded Markdown Web opens, edits, saves, and inserts a Host-picked image', async ({
  page,
}) => {
  await page.goto(MARKDOWN_WEB_HOST_URL)
  const officeFrame = page.frameLocator('#office-frame')

  await expect(page.locator('#host-state')).toHaveText('ready')
  await page.locator('#markdown-picker').setInputFiles({
    name: 'markdown-web.md',
    mimeType: 'text/markdown',
    buffer: Buffer.from('# Markdown Web\n\nInitial paragraph\n', 'utf8'),
  })

  await expect(officeFrame.locator('.doc-editor h1')).toHaveText('Markdown Web')
  const editor = officeFrame.locator('.doc-editor')
  const paragraph = editor.locator('p').filter({ hasText: 'Initial paragraph' })
  await expect(paragraph).toBeVisible()
  await paragraph.click()
  await editor.press('End')
  await editor.pressSequentially(' edited')
  await expect(paragraph).toHaveText('Initial paragraph edited')
  await expect(page.locator('#host-state')).toHaveText('dirty')

  await editor.press('Control+s')
  await expect(page.locator('#saved-text')).toContainText('Initial paragraph edited')
  await expect(page.locator('#host-state')).toHaveText('clean')

  const chooserPromise = page.waitForEvent('filechooser')
  await officeFrame.getByTitle('插入图片').click()
  const chooser = await chooserPromise
  await chooser.setFiles({
    name: 'uc-image.png',
    mimeType: 'image/png',
    buffer: ONE_PIXEL_PNG,
  })

  const image = editor.locator('img').last()
  await expect(image).toBeVisible()
  await expect(image).toHaveAttribute('src', /^data:image\/png;base64,/)
  await expect(page.locator('#host-state')).toHaveText('dirty')

  await editor.press('Control+s')
  await expect(page.locator('#saved-text')).toContainText('data:image/png;base64')
  await expect(page.locator('#host-state')).toHaveText('clean')
  await expect(officeFrame.locator('.ai-entry:visible')).toHaveCount(0)
})
