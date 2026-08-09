import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { expect, test, type Frame, type Page } from '@playwright/test'
import { addElement, createBlankPptx, openPptx, savePptx } from '@genoffice/pptx-engine'

const slidesWebUrl = process.env.SLIDES_WEB_E2E_URL
const hostUrl = process.env.SLIDES_WEB_HOST_E2E_URL
const fixturePath = resolve('fixtures/generated/slides-web-e2e.pptx')
const fixtureText = 'PPT Web 浏览器真实文件回归标题'

async function generateFixture(): Promise<void> {
  await mkdir(resolve('fixtures/generated'), { recursive: true })
  const blank = await createBlankPptx()
  const opened = await openPptx(blank)
  const slide = opened.deck.slides[0]
  if (!slide) throw new Error('Blank PPTX did not contain a slide')

  addElement(slide, {
    kind: 'textbox',
    offset: {
      x: 914400,
      y: 914400,
      cx: 7000000,
      cy: 1000000,
    },
    paragraphs: [{ runs: [{ text: fixtureText }] }],
  })

  await writeFile(fixturePath, await savePptx(opened))
}

async function slidesFrame(page: Page): Promise<Frame> {
  await expect
    .poll(() => page.frames().some((frame) => frame.url().includes(':5274/')), {
      timeout: 30_000,
    })
    .toBe(true)
  const frame = page.frames().find((candidate) => candidate.url().includes(':5274/'))
  if (!frame) throw new Error('Slides Web iframe was not found')
  return frame
}

async function slideJson(frame: Frame): Promise<string> {
  return frame.evaluate(async () => {
    const api = (window as any).slidesApi
    return JSON.stringify((await api.getRenderSlides()) ?? [])
  })
}

async function slideText(frame: Frame): Promise<string> {
  return frame.evaluate(async () => {
    const api = (window as any).slidesApi
    const slides = (await api.getRenderSlides()) ?? []
    return slides
      .flatMap((slide: any) => slide.nodes ?? [])
      .flatMap((node: any) => node.text?.lines ?? [])
      .flatMap((line: any) => line.runs ?? [])
      .map((run: any) => run.text ?? '')
      .join('')
  })
}

async function firstSourceId(frame: Frame): Promise<string> {
  const id = await frame.evaluate(async () => {
    const api = (window as any).slidesApi
    const slides = await api.getRenderSlides()
    return slides?.[0]?.nodes?.[0]?.sourceId ?? null
  })
  if (!id) throw new Error('No editable source element was found on the first slide')
  return id
}

async function replaceFirstElementText(frame: Frame, text: string): Promise<void> {
  const sourceId = await firstSourceId(frame)
  await frame.evaluate(
    async ({ sourceId: id, nextText }) => {
      const api = (window as any).slidesApi
      const result = await api.editText({
        slideIndex: 0,
        sourceId: id,
        paragraphs: [{ runs: [{ text: nextText }] }],
      })
      if (!result) throw new Error('Slides Web editText returned null')
    },
    { sourceId, nextText: text },
  )
}

async function addTextBox(frame: Frame, text: string): Promise<void> {
  await frame.evaluate(async (nextText) => {
    const api = (window as any).slidesApi
    const result = await api.addElement({
      slideIndex: 0,
      kind: 'textbox',
      xPx: 120,
      yPx: 100,
      wPx: 420,
      hPx: 80,
      fitWidthPx: 960,
      text: nextText,
    })
    if (!result) throw new Error('Slides Web addElement returned null')
  }, text)
}

async function clickFileCommand(frame: Frame, selector: string): Promise<void> {
  await frame.locator('.ribbon-tab-file').click()
  await frame.locator(selector).click()
}

test.beforeAll(async () => {
  await generateFixture()
})

test.describe('Slides Web', () => {
  test.skip(!slidesWebUrl, 'SLIDES_WEB_E2E_URL is only set by the Slides Web browser CI step')

  test('boots a productized standalone blank presentation', async ({ page }) => {
    const unexpectedDownloads: string[] = []
    page.on('download', (download) => unexpectedDownloads.push(download.suggestedFilename()))

    await page.goto(slidesWebUrl!)
    await expect
      .poll(
        async () =>
          JSON.parse(
            await page.evaluate(async () => {
              const api = (window as any).slidesApi
              return JSON.stringify((await api.getRenderSlides()) ?? [])
            }),
          ).length,
        { timeout: 30_000 },
      )
      .toBeGreaterThan(0)

    await expect(page.locator('.autosave-toggle')).toBeHidden()
    await expect(page.locator('.ai-dock')).toBeHidden()
    await expect(page.locator('.ai-rail')).toBeHidden()
    await expect(page.locator('text=GenOffice Slides Web failed to start')).toHaveCount(0)

    // Background activity must never turn a blank browser presentation into a download.
    await page.waitForTimeout(1_000)
    expect(unexpectedDownloads).toEqual([])
  })

  test('creates, opens, edits, saves, versions, exports, switches mode/locale, and exits through the iframe host', async ({
    page,
  }) => {
    test.skip(!hostUrl, 'SLIDES_WEB_HOST_E2E_URL is required for the iframe host flow')

    await page.goto(hostUrl!)
    const frame = await slidesFrame(page)

    // App Center flow: no selected node means office:new and a real blank PPTX.
    await expect
      .poll(async () => JSON.parse(await slideJson(frame)).length, { timeout: 30_000 })
      .toBe(1)
    await addTextBox(frame, 'App Center 新建演示文稿第一次保存')
    await expect(page.locator('#dirty-state')).toHaveText('dirty')

    page.once('dialog', async (dialog) => {
      expect(dialog.type()).toBe('prompt')
      await dialog.accept('PPT-Web-新建验证.pptx')
    })
    await page.locator('#save-button').click()
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#file-name')).toHaveText('PPT-Web-新建验证.pptx')
    await expect(page.locator('#dirty-state')).toHaveText('clean')

    // Open a real PPTX through the Host-owned file path. Render text may be split
    // into per-glyph runs, so compare the recombined rendered text rather than JSON.
    await page.locator('#pptx-picker').setInputFiles(fixturePath)
    await expect.poll(() => slideText(frame), { timeout: 30_000 }).toContain(fixtureText)

    // Embedded File menu matches the Word Web lifecycle contract.
    await frame.locator('.ribbon-tab-file').click()
    await expect(frame.locator('.file-menu-open')).toBeVisible()
    await expect(frame.locator('.file-menu-save')).toBeVisible()
    await expect(frame.locator('.file-menu-save-as')).toBeVisible()
    await expect(frame.locator('.file-menu-save-history')).toBeVisible()
    await expect(frame.locator('.file-menu-export-pptx')).toBeVisible()
    await expect(frame.locator('.file-menu-exit')).toBeVisible()
    await expect(frame.locator('.file-menu-electron-only').first()).toBeHidden()
    await frame.locator('.ribbon-tab-file').click()

    const editedText = 'PPT Web iframe 保存验证中文'
    await replaceFirstElementText(frame, editedText)
    await expect(page.locator('#dirty-state')).toHaveText('dirty')

    await page.locator('#save-button').click()
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#dirty-state')).toHaveText('clean')

    // The Host's current bytes must reparse as a real PPTX with the browser edit persisted.
    const savedDownloadPromise = page.waitForEvent('download')
    await page.locator('#download-button').click()
    const savedDownload = await savedDownloadPromise
    const savedPath = await savedDownload.path()
    expect(savedPath).toBeTruthy()
    const savedDeck = await openPptx(new Uint8Array(await readFile(savedPath!)))
    expect(JSON.stringify(savedDeck.deck.slides)).toContain(editedText)

    // Save As creates a new Host file identity/name.
    page.once('dialog', async (dialog) => {
      expect(dialog.type()).toBe('prompt')
      await dialog.accept('PPT-Web-另存为验证.pptx')
    })
    await clickFileCommand(frame, '.file-menu-save-as')
    await expect(page.locator('#host-state')).toHaveText('saved', { timeout: 30_000 })
    await expect(page.locator('#file-name')).toHaveText('PPT-Web-另存为验证.pptx')

    // History is a Host side effect and must not replace the current presentation.
    await clickFileCommand(frame, '.file-menu-save-history')
    await expect(page.locator('#host-state')).toHaveText('history:1', { timeout: 30_000 })
    await expect(page.locator('#file-name')).toHaveText('PPT-Web-另存为验证.pptx')

    // Export is a local PPTX download and also keeps current identity unchanged.
    const exportDownloadPromise = page.waitForEvent('download')
    await clickFileCommand(frame, '.file-menu-export-pptx')
    const exported = await exportDownloadPromise
    expect(exported.suggestedFilename()).toBe('PPT-Web-另存为验证.pptx')
    const exportedPath = await exported.path()
    const exportedDeck = await openPptx(new Uint8Array(await readFile(exportedPath!)))
    expect(JSON.stringify(exportedDeck.deck.slides)).toContain(editedText)
    await expect(page.locator('#file-name')).toHaveText('PPT-Web-另存为验证.pptx')

    // Host-authoritative preview/edit switching happens live without reloading the PPTX.
    await page.locator('#mode-button').click()
    await expect
      .poll(() => frame.evaluate(() => (window as any).slidesApi.getHostEditorMode()))
      .toBe('view')
    await expect(frame.locator('html')).toHaveClass(/office-view-mode/)
    await page.locator('#mode-button').click()
    await expect
      .poll(() => frame.evaluate(() => (window as any).slidesApi.getHostEditorMode()))
      .toBe('edit')

    // System language switching is propagated at runtime.
    await page.locator('#locale-button').click()
    await expect
      .poll(() => frame.evaluate(() => (window as any).slidesApi.getLanguage()))
      .toBe('en')
    await page.locator('#locale-button').click()
    await expect
      .poll(() => frame.evaluate(() => (window as any).slidesApi.getLanguage()))
      .toBe('zh')

    // Dirty Exit is guarded by GenOffice. Discard sends only the final Host close request.
    await replaceFirstElementText(frame, '退出前未保存修改')
    await expect(page.locator('#dirty-state')).toHaveText('dirty')
    await clickFileCommand(frame, '.file-menu-exit')
    await expect(frame.getByText('退出前保存更改？')).toBeVisible()
    await frame.getByRole('button', { name: '放弃并退出' }).click()
    await expect(page.locator('#host-state')).toHaveText('close requested')
  })
})
