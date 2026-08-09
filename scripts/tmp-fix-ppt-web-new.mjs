import { readFile, writeFile } from 'node:fs/promises'

async function replaceExact(path, before, after) {
  const source = await readFile(path, 'utf8')
  if (source.includes(after)) return false
  if (!source.includes(before)) throw new Error(`Expected block not found in ${path}`)
  await writeFile(path, source.replace(before, after))
  return true
}

await replaceExact(
  'apps/slides/src/web/product-policy.css',
  `html.office-web:not(.office-ai-enabled) .ai-dock,\nhtml.office-web:not(.office-ai-enabled) .ai-rail,`,
  `html.office-web:not(.office-ai-enabled) .ai-dock,\nhtml.office-web:not(.office-ai-enabled) .ai-rail,\nhtml.office-web:not(.office-ai-enabled) .stage-ai-bar,`,
)

await replaceExact(
  'packages/pptx-engine/src/builtin-layouts.ts',
  `  archive.entries.set(\n    masterRelsPath,\n    Buffer.from(\n      masterRels.replace(\n        '</Relationships>',\n        \`<Relationship Id="\${rid}" Type="\${LAYOUT_REL_TYPE}" Target="../slideLayouts/\${layoutPath.slice('ppt/slideLayouts/'.length)}"/></Relationships>\`,\n      ),\n      'utf8',\n    ),\n  )`,
  `  archive.entries.set(\n    masterRelsPath,\n    encodeUtf8(\n      masterRels.replace(\n        '</Relationships>',\n        \`<Relationship Id="\${rid}" Type="\${LAYOUT_REL_TYPE}" Target="../slideLayouts/\${layoutPath.slice('ppt/slideLayouts/'.length)}"/></Relationships>\`,\n      ),\n    ),\n  )`,
)

await replaceExact(
  'packages/pptx-engine/src/builtin-layouts.ts',
  `  archive.entries.set(masterPath, Buffer.from(nextMaster))`,
  `  archive.entries.set(masterPath, encodeUtf8(nextMaster))`,
)

await replaceExact(
  'apps/slides/src/web/slides-api.ts',
  `    newBlank: async (fitWidthPx: number) => {\n      const bytes = await createBlankPptx()\n      return replaceOpenedFromBytes(bytesToArrayBuffer(bytes), fitWidthPx, null)\n    },`,
  `    newBlank: async (fitWidthPx: number) => {\n      const bytes = await createBlankPptx()\n      const result = await replaceOpenedFromBytes(bytesToArrayBuffer(bytes), fitWidthPx, null)\n      if (!session) return result\n\n      // PowerPoint starts a new presentation with an editable Title Slide rather\n      // than a physically empty canvas. Keep createBlankPptx() unchanged for the\n      // desktop/engine callers, and apply the standard layout only at the Web\n      // product boundary. The placeholders stay true OOXML placeholders and save\n      // normally once the Host assigns the first real file identity.\n      const titleLayoutPath = ensureBuiltinLayout(\n        session.opened.archive,\n        session.opened.deck.size,\n        'titleSlide',\n      )\n      if (!titleLayoutPath) return result\n      if (!setSlideLayout(session.opened, 0, titleLayoutPath)) return result\n\n      // Applying the default template is initialization, not a user edit.\n      setDirtyState(false)\n      return openResult() ?? result\n    },`,
)

await replaceExact(
  'apps/slides/tests/web-slides-api.test.ts',
  `    const blank = await controller.slidesApi.newBlank(960)\n    expect(blank.slides).toHaveLength(1)\n\n    const added = await controller.slidesApi.addElement({\n      slideIndex: 0,\n      kind: 'textbox',\n      xPx: 100,\n      yPx: 100,\n      wPx: 400,\n      hPx: 80,\n      fitWidthPx: 960,\n      text: '新建 PPT Web',\n    })\n    expect(added).not.toBeNull()`,
  `    const blank = await controller.slidesApi.newBlank(960)\n    expect(blank.slides).toHaveLength(1)\n    const editablePlaceholder = blank.slides[0]?.nodes.find(\n      (node: any) => node.placeholder && (node.type === 'text' || node.type === 'shape'),\n    )\n    expect(editablePlaceholder).toBeTruthy()\n\n    const added = await controller.slidesApi.editText({\n      slideIndex: 0,\n      sourceId: editablePlaceholder!.sourceId,\n      paragraphs: [{ runs: [{ text: '新建 PPT Web' }] }],\n    })\n    expect(added).not.toBeNull()`,
)

await replaceExact(
  'e2e/slides-web.spec.ts',
  `    await expect(page.locator('.ai-rail')).toBeHidden()\n    await expect(page.locator('text=GenOffice Slides Web failed to start')).toHaveCount(0)`,
  `    await expect(page.locator('.ai-rail')).toBeHidden()\n    await expect(page.locator('.stage-ai-bar')).toBeHidden()\n    await expect(page.locator('text=GenOffice Slides Web failed to start')).toHaveCount(0)`,
)

await replaceExact(
  'e2e/slides-web.spec.ts',
  `    await addTextBox(frame, 'App Center 新建演示文稿第一次保存')`,
  `    // A new PPT now contains standard title placeholders and is immediately editable.\n    await replaceFirstElementText(frame, 'App Center 新建演示文稿第一次保存')`,
)

console.log('PPT Web new-presentation / AI policy patch applied.')
