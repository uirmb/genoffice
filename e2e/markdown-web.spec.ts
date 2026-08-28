import { expect, test } from '@playwright/test'

const MARKDOWN_WEB_URL = process.env.MARKDOWN_WEB_E2E_URL ?? 'http://127.0.0.1:5277/'
const MARKDOWN_WEB_HOST_URL = process.env.MARKDOWN_WEB_HOST_E2E_URL ?? 'http://127.0.0.1:8084/'

const ONE_PIXEL_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2n0AAAAAASUVORK5CYII=',
  'base64',
)

async function requestHostWindowClose(page: import('@playwright/test').Page, requestId: string) {
  await page.locator('#office-frame').evaluate((element, id) => {
    const frame = element as HTMLIFrameElement
    frame.contentWindow?.postMessage(
      {
        protocol: 1,
        type: 'office:request-close',
        requestId: id,
        payload: { reason: 'window-close' },
      },
      new URL(frame.src).origin,
    )
  }, requestId)
}

test('standalone Markdown Web starts as an editable AI-free document', async ({ page }) => {
  await page.goto(MARKDOWN_WEB_URL)

  await expect(page.locator('.doc-editor')).toBeVisible()
  await expect(page.locator('.doc-editor')).toHaveAttribute('contenteditable', 'true')
  await expect(page.locator('.markdown-file-tab')).toBeVisible()
  await expect(page.locator('.ai-entry:visible')).toHaveCount(0)
  await expect(page.locator('.ai-dock:visible')).toHaveCount(0)
})

test('embedded Markdown Web uses the iframe locale on first render and follows live Host changes', async ({
  page,
}) => {
  await page.goto(MARKDOWN_WEB_HOST_URL)
  const frame = page.locator('#office-frame')
  await expect(page.locator('#host-state')).toHaveText('ready')

  // UC includes locale on the plugin iframe URL. Reload the demo iframe with the
  // same host origin plus an English locale before any document is bound.
  await frame.evaluate((element) => {
    const iframe = element as HTMLIFrameElement
    const url = new URL(iframe.src)
    url.searchParams.set('locale', 'en-US')
    iframe.src = url.toString()
  })

  const officeFrame = page.frameLocator('#office-frame')
  await expect(officeFrame.locator('.markdown-file-tab')).toHaveText('File')

  await frame.evaluate((element) => {
    const iframe = element as HTMLIFrameElement
    iframe.contentWindow?.postMessage(
      {
        protocol: 1,
        type: 'office:set-locale',
        payload: { locale: 'zh-CN' },
      },
      new URL(iframe.src).origin,
    )
  })
  await expect(officeFrame.locator('.markdown-file-tab')).toHaveText('文件')
})

test('embedded Markdown Web File menu, save UI, exit prompt, Host close, and asset flow match UC lifecycle', async ({
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
  const fileTab = officeFrame.locator('.markdown-file-tab')

  await expect(fileTab).toHaveText('文件')
  await fileTab.click()
  const fileMenu = officeFrame.locator('.markdown-file-menu')
  await expect(fileMenu).toBeVisible()
  await expect(fileMenu.getByRole('menuitem', { name: /存为新的历史版本/ })).toBeVisible()
  await expect(fileMenu.getByRole('menuitem', { name: '下载到本地' })).toBeVisible()
  await expect(fileMenu.getByRole('menuitem', { name: '退出' })).toBeVisible()
  await expect(fileMenu.locator('.markdown-file-menu-separator')).toHaveCount(0)
  await expect(fileMenu.locator('button').filter({ hasText: '保存' })).toBeDisabled()

  const quickSave = officeFrame.locator('.markdown-quick-save')
  await expect(quickSave).toBeDisabled()
  await expect(quickSave).toHaveCSS('width', '30px')
  await expect(quickSave).toHaveCSS('height', '28px')
  await expect(quickSave).toHaveCSS('border-radius', '4px')
  await expect(officeFrame.locator('.markdown-quick-save + .rb-sep')).toBeHidden()
  await fileTab.click()

  const paragraph = editor.locator('p').filter({ hasText: 'Initial paragraph' })
  await expect(paragraph).toBeVisible()
  await paragraph.click()
  await editor.press('End')
  await editor.pressSequentially(' edited')
  await expect(paragraph).toHaveText('Initial paragraph edited')
  await expect(page.locator('#host-state')).toHaveText('dirty')

  await expect(quickSave).toBeEnabled()
  await quickSave.click()
  await expect(page.locator('#saved-text')).toContainText('Initial paragraph edited')
  await expect(page.locator('#host-state')).toHaveText('clean')
  await expect(quickSave).toBeDisabled()

  const openChooserPromise = page.waitForEvent('filechooser')
  await fileTab.click()
  await fileMenu.locator('button').filter({ hasText: '打开' }).click()
  const openChooser = await openChooserPromise
  await openChooser.setFiles({
    name: 'second.md',
    mimeType: 'text/markdown',
    buffer: Buffer.from('# Second Markdown\n\nOpened from File menu\n', 'utf8'),
  })

  await expect(editor.locator('h1')).toHaveText('Second Markdown')
  const secondParagraph = editor.locator('p').filter({ hasText: 'Opened from File menu' })
  await expect(secondParagraph).toBeVisible()
  await expect(page.locator('#file-name')).toHaveText('second.md')
  await expect(page.locator('#host-state')).toHaveText('clean')

  await fileTab.click()
  await fileMenu.locator('.markdown-file-menu-history').click()
  await expect(page.locator('#saved-text')).toContainText('Second Markdown')
  await expect(page.locator('#host-state')).toHaveText('clean')

  await fileTab.click()
  await fileMenu.locator('.markdown-file-menu-download').click()
  await expect(page.locator('#host-state')).toHaveText('downloaded')

  const assetChooserPromise = page.waitForEvent('filechooser')
  await officeFrame.getByTitle('插入图片').click()
  const assetChooser = await assetChooserPromise
  await assetChooser.setFiles({
    name: 'uc-image.png',
    mimeType: 'image/png',
    buffer: ONE_PIXEL_PNG,
  })

  const image = editor.locator('img').last()
  await expect(image).toBeVisible()
  await expect(image).toHaveAttribute('src', /^data:image\/png;base64,/)
  await expect(page.locator('#host-state')).toHaveText('dirty')

  await fileTab.click()
  await fileMenu.locator('button').filter({ hasText: '另存为' }).click()
  await expect(page.locator('#saved-text')).toContainText('data:image/png;base64')
  await expect(page.locator('#host-state')).toHaveText('clean')
  await expect(officeFrame.locator('.ai-entry:visible')).toHaveCount(0)

  // File -> Exit: dirty documents use the same Word-style three-choice in-page prompt.
  await secondParagraph.click()
  await editor.press('End')
  await editor.pressSequentially(' file-exit')
  await expect(page.locator('#host-state')).toHaveText('dirty')

  await fileTab.click()
  await fileMenu.locator('.markdown-file-menu-exit').click()
  const exitModal = officeFrame.locator('.modal-backdrop')
  await expect(exitModal).toBeVisible()
  await expect(exitModal.getByRole('heading')).toHaveText('退出前保存更改?')
  await expect(exitModal).toContainText('此文档有尚未保存的更改。')
  await expect(exitModal.getByRole('button', { name: '取消' })).toBeVisible()
  await expect(exitModal.getByRole('button', { name: '放弃并退出' })).toBeVisible()
  await expect(exitModal.getByRole('button', { name: '保存并退出' })).toBeVisible()
  await exitModal.getByRole('button', { name: '保存并退出' }).click()
  await expect(page.locator('#saved-text')).toContainText('file-exit')
  await expect(page.locator('#host-state')).toHaveText('close approved (file-menu)')
  await expect(exitModal).toHaveCount(0)

  // Host/window close uses the identical prompt and completes the existing
  // office:request-close -> approve/cancel bridge rather than auto-saving.
  await secondParagraph.click()
  await editor.press('End')
  await editor.pressSequentially(' window-exit')
  await expect(page.locator('#host-state')).toHaveText('dirty')

  await requestHostWindowClose(page, 'e2e-window-close-cancel')
  await expect(exitModal).toBeVisible()
  await exitModal.getByRole('button', { name: '取消' }).click()
  await expect(page.locator('#host-state')).toHaveText('close cancelled')
  await expect(exitModal).toHaveCount(0)

  await requestHostWindowClose(page, 'e2e-window-close-discard')
  await expect(exitModal).toBeVisible()
  await exitModal.getByRole('button', { name: '放弃并退出' }).click()
  await expect(page.locator('#host-state')).toHaveText('close approved (window-close)')
  await expect(exitModal).toHaveCount(0)
})
