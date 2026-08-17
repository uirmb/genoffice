from pathlib import Path


def replace_once(path: str, old: str, new: str) -> None:
    target = Path(path)
    text = target.read_text()
    count = text.count(old)
    if count != 1:
        raise SystemExit(f'{path}: expected one match, found {count}: {old[:80]!r}')
    target.write_text(text.replace(old, new, 1))


replace_once(
    'packages/office-host-api/src/index.ts',
    "export type OfficeExportFormat = 'docx' | 'pptx' | 'xlsx'",
    "export type OfficeExportFormat = 'docx' | 'pptx' | 'xlsx' | 'markdown'",
)

replace_once(
    'apps/markdown/src/shared/ipc.ts',
    "export type SaveMarkdownResult =\n  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }\n",
    "export type SaveMarkdownResult =\n  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }\n\nexport interface MarkdownTextRequest {\n  /** Full UTF-8 Markdown text including frontmatter. */\n  text: string\n}\n\nexport type MarkdownHostActionResult = { ok: true } | { ok: false; error: string }\n",
)
replace_once(
    'apps/markdown/src/shared/ipc.ts',
    "  save(request: SaveMarkdownRequest): Promise<SaveMarkdownResult>\n  /** Mirror unsaved-changes state to the main process; drives the save prompt before closing a tab/window */",
    "  save(request: SaveMarkdownRequest): Promise<SaveMarkdownResult>\n  /** Web-only UC Host actions. Desktop preload may omit these hooks. */\n  saveHistoryVersion?(request: MarkdownTextRequest): Promise<MarkdownHostActionResult>\n  download?(request: MarkdownTextRequest): Promise<MarkdownHostActionResult>\n  exit?(): Promise<void>\n  /** Mirror unsaved-changes state to the main process; drives the save prompt before closing a tab/window */",
)

replace_once(
    'apps/markdown/src/web/markdown-api.ts',
    "  OpenMarkdownResult,\n  SaveMarkdownRequest,\n  SaveMarkdownResult,\n  SaveMode,",
    "  OpenMarkdownResult,\n  MarkdownHostActionResult,\n  MarkdownTextRequest,\n  SaveMarkdownRequest,\n  SaveMarkdownResult,\n  SaveMode,",
)
replace_once(
    'apps/markdown/src/web/markdown-api.ts',
    "  setDirty(dirty: boolean): void {\n    this.dirty = dirty\n    this.host.setDirty(dirty)\n  }",
    "  async saveHistoryVersion(request: MarkdownTextRequest): Promise<MarkdownHostActionResult> {\n    const existing = this.currentFile\n    if (!existing) return { ok: false, error: 'Save the Markdown document before creating history.' }\n    if (!this.host.saveHistoryVersion) {\n      return { ok: false, error: 'History versions are not supported by this host.' }\n    }\n\n    const bytes = new TextEncoder().encode(request.text).buffer\n    const result = await this.host.saveHistoryVersion({\n      file: descriptorOf(existing),\n      bytes,\n      baseVersion: existing.version,\n    })\n    if (!result.ok) {\n      return { ok: false, error: result.error || 'Creating a Markdown history version failed.' }\n    }\n\n    const saved = result.file ?? { ...descriptorOf(existing), size: bytes.byteLength }\n    this.currentFile = {\n      ...saved,\n      mimeType: saved.mimeType || 'text/markdown',\n      size: bytes.byteLength,\n      bytes: bytes.slice(0),\n      transport: 'buffer',\n    }\n    this.host.setTitle(this.currentFile.name)\n    return { ok: true }\n  }\n\n  async download(request: MarkdownTextRequest): Promise<MarkdownHostActionResult> {\n    const download = this.host.downloadDocument ?? this.host.exportDocument\n    if (!download) return { ok: false, error: 'Download is not supported by this host.' }\n\n    const bytes = new TextEncoder().encode(request.text).buffer\n    const existing = this.currentFile\n    const name = ensureMarkdownName(existing?.name || 'Untitled.md')\n    const file: OfficeFileDescriptor = existing\n      ? { ...descriptorOf(existing), name, size: bytes.byteLength }\n      : {\n          id: 'download:markdown',\n          name,\n          mimeType: 'text/markdown',\n          size: bytes.byteLength,\n          version: null,\n          transport: 'buffer',\n        }\n    const result = await download.call(this.host, {\n      format: 'markdown',\n      file,\n      bytes,\n    })\n    return result.ok\n      ? { ok: true }\n      : { ok: false, error: result.error || 'Downloading Markdown failed.' }\n  }\n\n  async exit(): Promise<void> {\n    if (this.host.approveClose) await this.host.approveClose()\n    else await this.host.requestClose?.()\n  }\n\n  setDirty(dirty: boolean): void {\n    this.dirty = dirty\n    this.host.setDirty(dirty)\n  }",
)

icons = Path('apps/markdown/src/renderer/components/icons.tsx')
icons_text = icons.read_text()
if 'export function IconSave(' not in icons_text:
    icons.write_text(
        icons_text
        + "\n\nexport function IconSave(props: IconProps) {\n  return (\n    <Svg {...props}>\n      <path d=\"M3 2.5h8.4L13.5 4.6V13.5H2.5V2.5z\" />\n      <path d=\"M5 2.5v4h5v-4\" />\n      <path d=\"M5 10h6v3.5H5z\" />\n    </Svg>\n  )\n}\n"
    )

Path('apps/markdown/src/renderer/components/FileMenu.tsx').write_text("""import { useEffect, useRef, useState } from 'react'
import type { Lang } from '@genoffice/i18n'
import { useI18n } from '../i18n/locale'

interface Props {
  disabled: boolean
  canSave: boolean
  canSaveHistoryVersion: boolean
  canDownload: boolean
  canExit: boolean
  onOpen: () => void
  onSave: () => void
  onSaveAs: () => void
  onSaveHistoryVersion: () => void
  onDownload: () => void
  onExit: () => void
}

type FileLabels = {
  file: string
  open: string
  save: string
  saveAs: string
  saveHistoryVersion: string
  download: string
  exit: string
}

const BASE_LABELS: Record<Lang, Pick<FileLabels, 'file' | 'open' | 'save' | 'saveAs'>> = {
  zh: { file: '文件', open: '打开', save: '保存', saveAs: '另存为' },
  'zh-TW': { file: '檔案', open: '開啟', save: '儲存', saveAs: '另存新檔' },
  en: { file: 'File', open: 'Open', save: 'Save', saveAs: 'Save As' },
  ja: { file: 'ファイル', open: '開く', save: '保存', saveAs: '名前を付けて保存' },
  ko: { file: '파일', open: '열기', save: '저장', saveAs: '다른 이름으로 저장' },
  fr: { file: 'Fichier', open: 'Ouvrir', save: 'Enregistrer', saveAs: 'Enregistrer sous' },
  de: { file: 'Datei', open: 'Öffnen', save: 'Speichern', saveAs: 'Speichern unter' },
  es: { file: 'Archivo', open: 'Abrir', save: 'Guardar', saveAs: 'Guardar como' },
  th: { file: 'ไฟล์', open: 'เปิด', save: 'บันทึก', saveAs: 'บันทึกเป็น' },
  id: { file: 'File', open: 'Buka', save: 'Simpan', saveAs: 'Simpan sebagai' },
  ru: { file: 'Файл', open: 'Открыть', save: 'Сохранить', saveAs: 'Сохранить как' },
  ar: { file: 'ملف', open: 'فتح', save: 'حفظ', saveAs: 'حفظ باسم' },
  pt: { file: 'Arquivo', open: 'Abrir', save: 'Salvar', saveAs: 'Salvar como' },
  it: { file: 'File', open: 'Apri', save: 'Salva', saveAs: 'Salva con nome' },
  pl: { file: 'Plik', open: 'Otwórz', save: 'Zapisz', saveAs: 'Zapisz jako' },
  nl: { file: 'Bestand', open: 'Openen', save: 'Opslaan', saveAs: 'Opslaan als' },
  ms: { file: 'Fail', open: 'Buka', save: 'Simpan', saveAs: 'Simpan sebagai' },
  he: { file: 'קובץ', open: 'פתיחה', save: 'שמירה', saveAs: 'שמירה בשם' },
  hi: { file: 'फ़ाइल', open: 'खोलें', save: 'सहेजें', saveAs: 'इस रूप में सहेजें' },
}

const ACTION_LABELS: Record<Lang, Pick<FileLabels, 'saveHistoryVersion' | 'download' | 'exit'>> = {
  zh: { saveHistoryVersion: '存为新的历史版本', download: '下载到本地', exit: '退出' },
  'zh-TW': { saveHistoryVersion: '儲存為新的歷史版本', download: '下載到本機', exit: '結束' },
  en: { saveHistoryVersion: 'Save as New History Version', download: 'Download', exit: 'Exit' },
  ja: { saveHistoryVersion: '新しい履歴版として保存', download: 'ダウンロード', exit: '終了' },
  ko: { saveHistoryVersion: '새 기록 버전으로 저장', download: '다운로드', exit: '종료' },
  fr: { saveHistoryVersion: 'Enregistrer une nouvelle version', download: 'Télécharger', exit: 'Quitter' },
  de: { saveHistoryVersion: 'Als neue Verlaufsversion speichern', download: 'Herunterladen', exit: 'Beenden' },
  es: { saveHistoryVersion: 'Guardar nueva versión del historial', download: 'Descargar', exit: 'Salir' },
  th: { saveHistoryVersion: 'บันทึกเป็นเวอร์ชันประวัติใหม่', download: 'ดาวน์โหลด', exit: 'ออก' },
  id: { saveHistoryVersion: 'Simpan sebagai versi riwayat baru', download: 'Unduh', exit: 'Keluar' },
  ru: { saveHistoryVersion: 'Сохранить новую версию истории', download: 'Скачать', exit: 'Выйти' },
  ar: { saveHistoryVersion: 'حفظ كإصدار سجل جديد', download: 'تنزيل', exit: 'خروج' },
  pt: { saveHistoryVersion: 'Salvar nova versão do histórico', download: 'Baixar', exit: 'Sair' },
  it: { saveHistoryVersion: 'Salva nuova versione cronologia', download: 'Scarica', exit: 'Esci' },
  pl: { saveHistoryVersion: 'Zapisz nową wersję historii', download: 'Pobierz', exit: 'Wyjdź' },
  nl: { saveHistoryVersion: 'Opslaan als nieuwe geschiedenisversie', download: 'Downloaden', exit: 'Afsluiten' },
  ms: { saveHistoryVersion: 'Simpan sebagai versi sejarah baharu', download: 'Muat turun', exit: 'Keluar' },
  he: { saveHistoryVersion: 'שמירה כגרסת היסטוריה חדשה', download: 'הורדה', exit: 'יציאה' },
  hi: { saveHistoryVersion: 'नए इतिहास संस्करण के रूप में सहेजें', download: 'डाउनलोड', exit: 'बाहर निकलें' },
}

export function FileMenu({
  disabled,
  canSave,
  canSaveHistoryVersion,
  canDownload,
  canExit,
  onOpen,
  onSave,
  onSaveAs,
  onSaveHistoryVersion,
  onDownload,
  onExit,
}: Props) {
  const { lang } = useI18n()
  const labels: FileLabels = { ...BASE_LABELS[lang], ...ACTION_LABELS[lang] }
  const [open, setOpen] = useState(false)
  const rootRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const onPointerDown = (event: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(event.target as Node)) setOpen(false)
    }
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setOpen(false)
    }
    window.addEventListener('mousedown', onPointerDown)
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('mousedown', onPointerDown)
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [])

  const run = (action: () => void) => {
    setOpen(false)
    action()
  }

  return (
    <div className="markdown-file-tab-wrap" ref={rootRef}>
      <button
        type="button"
        className={`markdown-file-tab${open ? ' open' : ''}`}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => setOpen((value) => !value)}
      >
        {labels.file}
      </button>
      {open && (
        <div className="markdown-file-menu" role="menu">
          <button type="button" role="menuitem" onClick={() => run(onOpen)}>
            <span>{labels.open}</span>
            <span className="markdown-file-menu-key">Ctrl+O</span>
          </button>
          <button type="button" role="menuitem" disabled={!canSave} onClick={() => run(onSave)}>
            <span>{labels.save}</span>
            <span className="markdown-file-menu-key">Ctrl+S</span>
          </button>
          <button type="button" role="menuitem" onClick={() => run(onSaveAs)}>
            <span>{labels.saveAs}</span>
            <span className="markdown-file-menu-key">Ctrl+Shift+S</span>
          </button>
          <div className="markdown-file-menu-separator" role="separator" />
          <button
            type="button"
            role="menuitem"
            className="markdown-file-menu-history"
            disabled={!canSaveHistoryVersion}
            onClick={() => run(onSaveHistoryVersion)}
          >
            <span>{labels.saveHistoryVersion}</span>
          </button>
          <button
            type="button"
            role="menuitem"
            className="markdown-file-menu-download"
            disabled={!canDownload}
            onClick={() => run(onDownload)}
          >
            <span>{labels.download}</span>
          </button>
          <div className="markdown-file-menu-separator" role="separator" />
          <button
            type="button"
            role="menuitem"
            className="markdown-file-menu-exit"
            disabled={!canExit}
            onClick={() => run(onExit)}
          >
            <span>{labels.exit}</span>
          </button>
        </div>
      )}
    </div>
  )
}
""")

replace_once(
    'apps/markdown/src/renderer/App.tsx',
    "import { FileMenu } from './components/FileMenu'\n",
    "import { FileMenu } from './components/FileMenu'\nimport { IconSave } from './components/icons'\n",
)
replace_once(
    'apps/markdown/src/renderer/App.tsx',
    "  const openDocument = useCallback(async (): Promise<void> => {",
    """  const saveHistoryVersion = useCallback(async (): Promise<boolean> => {
    const action = window.markdownApi.saveHistoryVersion
    const current = editorRef.current
    if (
      !action ||
      !current ||
      !filePathRef.current ||
      statusRef.current !== 'ready' ||
      savingRef.current
    ) {
      return false
    }

    savingRef.current = true
    setSaveState('saving')
    try {
      const docAtSave = current.state.doc
      const fmAtSave = envelopeRef.current.frontmatter
      const text = serializeDocText(envelopeRef.current, current.getMarkdown())
      const result = await action.call(window.markdownApi, { text })
      if (!result.ok) {
        console.error('[markdown] save history failed:', result.error)
        setSaveState('failed')
        return false
      }

      const unchanged =
        editorRef.current?.state.doc === docAtSave && envelopeRef.current.frontmatter === fmAtSave
      if (unchanged) {
        dirtyRef.current = false
        setDirty(false)
        window.markdownApi.setDirty(false)
        setSaveState('saved')
      } else {
        dirtyRef.current = true
        setDirty(true)
        window.markdownApi.setDirty(true)
        setSaveState('idle')
      }
      return true
    } catch (error) {
      console.error('[markdown] save history failed:', error)
      setSaveState('failed')
      return false
    } finally {
      savingRef.current = false
    }
  }, [])

  const downloadMarkdown = useCallback(async (): Promise<void> => {
    const action = window.markdownApi.download
    const current = editorRef.current
    if (!action || !current || statusRef.current !== 'ready' || savingRef.current) return
    const text = serializeDocText(envelopeRef.current, current.getMarkdown())
    try {
      const result = await action.call(window.markdownApi, { text })
      if (!result.ok) console.error('[markdown] download failed:', result.error)
    } catch (error) {
      console.error('[markdown] download failed:', error)
    }
  }, [])

  const exitMarkdown = useCallback(async (): Promise<void> => {
    const exit = window.markdownApi.exit
    if (!exit || statusRef.current !== 'ready' || savingRef.current) return

    if (dirtyRef.current) {
      const isChinese = document.documentElement.lang.toLowerCase().startsWith('zh')
      const saveFirst = window.confirm(
        isChinese
          ? '当前文档有未保存的更改。退出前需要先保存，是否保存并退出？'
          : 'This document has unsaved changes. Save before exiting?',
      )
      if (!saveFirst) return
      const saved = await doSave('save')
      if (!saved) return
    }

    await exit.call(window.markdownApi)
  }, [doSave])

  const openDocument = useCallback(async (): Promise<void> => {""",
)
replace_once(
    'apps/markdown/src/renderer/App.tsx',
    """        leading={
          window.markdownApi.openDocument ? (
            <FileMenu
              disabled={status !== 'ready' || saveState === 'saving'}
              canSave={status === 'ready' && dirty && saveState !== 'saving'}
              onOpen={() => void openDocument()}
              onSave={() => void doSave('save')}
              onSaveAs={() => void doSave('saveAs')}
            />
          ) : undefined
        }""",
    """        leading={
          window.markdownApi.openDocument ? (
            <>
              <FileMenu
                disabled={status !== 'ready' || saveState === 'saving'}
                canSave={status === 'ready' && dirty && saveState !== 'saving'}
                canSaveHistoryVersion={
                  status === 'ready' && Boolean(filePath) && saveState !== 'saving'
                }
                canDownload={status === 'ready' && saveState !== 'saving'}
                canExit={status === 'ready' && saveState !== 'saving'}
                onOpen={() => void openDocument()}
                onSave={() => void doSave('save')}
                onSaveAs={() => void doSave('saveAs')}
                onSaveHistoryVersion={() => void saveHistoryVersion()}
                onDownload={() => void downloadMarkdown()}
                onExit={() => void exitMarkdown()}
              />
              <button
                type="button"
                className="rb-btn markdown-quick-save"
                title={document.documentElement.lang.toLowerCase().startsWith('zh') ? '保存' : 'Save'}
                aria-label={document.documentElement.lang.toLowerCase().startsWith('zh') ? '保存' : 'Save'}
                disabled={status !== 'ready' || !dirty || saveState === 'saving'}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => void doSave('save')}
              >
                <IconSave size={16} />
              </button>
            </>
          ) : undefined
        }""",
)

replace_once(
    'apps/markdown/src/web/product-policy.css',
    "html.genoffice-markdown-web .markdown-file-menu-key {\n  margin-left: 24px;\n  color: #605e5c;\n  font-size: 12px;\n  white-space: nowrap;\n}\n",
    "html.genoffice-markdown-web .markdown-file-menu-key {\n  margin-left: 24px;\n  color: #605e5c;\n  font-size: 12px;\n  white-space: nowrap;\n}\n\nhtml.genoffice-markdown-web .markdown-file-menu-separator {\n  height: 1px;\n  margin: 4px 8px;\n  background: #edebe9;\n}\n\nhtml.genoffice-markdown-web .markdown-quick-save {\n  flex: 0 0 auto;\n}\n",
)

replace_once(
    'examples/web-markdown-host/src/main.ts',
    "    saveHistoryVersion: false,",
    "    saveHistoryVersion: true,",
)
replace_once(
    'examples/web-markdown-host/src/main.ts',
    "    download: false,",
    "    download: true,",
)
replace_once(
    'examples/web-markdown-host/src/main.ts',
    "    case 'office:pick-assets':\n      pendingAssetRequestId = message.requestId\n      assetPicker.accept = message.payload.accept?.join(',') || 'image/png,image/jpeg,image/gif'\n      assetPicker.click()\n      return",
    """    case 'office:save-history-version': {
      saveVersion += 1
      const bytes = message.payload.bytes.slice(0)
      const savedDescriptor: OfficeFileDescriptor = {
        ...message.payload.file,
        size: bytes.byteLength,
        version: `history-${saveVersion}`,
        transport: 'buffer',
      }
      currentFile = { ...savedDescriptor, bytes, transport: 'buffer' }
      savedText.textContent =
        new TextDecoder().decode(bytes).replace(/\\s+/g, ' ').trim() || '(empty)'
      render()
      setHostState('history saved')
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-history-version-result',
        requestId: message.requestId,
        payload: { ok: true, file: savedDescriptor },
      })
      return
    }
    case 'office:download-document':
      if (message.payload.format !== 'markdown') {
        send({
          protocol: OFFICE_PROTOCOL_VERSION,
          type: 'office:download-document-result',
          requestId: message.requestId,
          payload: { ok: false, code: 'UNSUPPORTED_FORMAT', error: 'Expected markdown.' },
        })
        return
      }
      setHostState('downloaded')
      send({
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:download-document-result',
        requestId: message.requestId,
        payload: { ok: true },
      })
      return
    case 'office:pick-assets':
      pendingAssetRequestId = message.requestId
      assetPicker.accept = message.payload.accept?.join(',') || 'image/png,image/jpeg,image/gif'
      assetPicker.click()
      return""",
)

replace_once(
    'e2e/markdown-web.spec.ts',
    "  await expect(fileMenu).toBeVisible()\n  await expect(fileMenu.locator('button').filter({ hasText: '保存' })).toBeDisabled()\n  await fileTab.click()",
    "  await expect(fileMenu).toBeVisible()\n  await expect(fileMenu.getByRole('button', { name: /存为新的历史版本/ })).toBeVisible()\n  await expect(fileMenu.getByRole('button', { name: '下载到本地' })).toBeVisible()\n  await expect(fileMenu.getByRole('button', { name: '退出' })).toBeVisible()\n  await expect(fileMenu.locator('button').filter({ hasText: '保存' })).toBeDisabled()\n  const quickSave = officeFrame.locator('.markdown-quick-save')\n  await expect(quickSave).toBeDisabled()\n  await fileTab.click()",
)
replace_once(
    'e2e/markdown-web.spec.ts',
    "  await fileTab.click()\n  const saveButton = fileMenu.locator('button').filter({ hasText: '保存' })\n  await expect(saveButton).toBeEnabled()\n  await saveButton.click()\n  await expect(page.locator('#saved-text')).toContainText('Initial paragraph edited')\n  await expect(page.locator('#host-state')).toHaveText('clean')",
    "  await expect(quickSave).toBeEnabled()\n  await quickSave.click()\n  await expect(page.locator('#saved-text')).toContainText('Initial paragraph edited')\n  await expect(page.locator('#host-state')).toHaveText('clean')\n  await expect(quickSave).toBeDisabled()",
)
replace_once(
    'e2e/markdown-web.spec.ts',
    "  await expect(page.locator('#host-state')).toHaveText('clean')\n\n  const assetChooserPromise = page.waitForEvent('filechooser')",
    "  await expect(page.locator('#host-state')).toHaveText('clean')\n\n  await fileTab.click()\n  await fileMenu.locator('.markdown-file-menu-history').click()\n  await expect(page.locator('#host-state')).toHaveText('history saved')\n\n  await fileTab.click()\n  await fileMenu.locator('.markdown-file-menu-download').click()\n  await expect(page.locator('#host-state')).toHaveText('downloaded')\n\n  const assetChooserPromise = page.waitForEvent('filechooser')",
)
replace_once(
    'e2e/markdown-web.spec.ts',
    "  await expect(page.locator('#host-state')).toHaveText('clean')\n  await expect(officeFrame.locator('.ai-entry:visible')).toHaveCount(0)\n})",
    "  await expect(page.locator('#host-state')).toHaveText('clean')\n  await expect(officeFrame.locator('.ai-entry:visible')).toHaveCount(0)\n\n  await fileTab.click()\n  await fileMenu.locator('.markdown-file-menu-exit').click()\n  await expect(page.locator('#host-state')).toHaveText('close approved (file-menu)')\n})",
)
