from pathlib import Path
import re


def read(path: str) -> str:
    return Path(path).read_text()


def write(path: str, text: str) -> None:
    Path(path).write_text(text)


def rep(path: str, old: str, new: str, count: int = 1) -> None:
    text = read(path)
    if old not in text:
        raise SystemExit(f"anchor not found in {path}: {old[:100]!r}")
    write(path, text.replace(old, new, count))


# 1) Shared host API ---------------------------------------------------------
path = "packages/office-host-api/src/index.ts"
rep(
    path,
    "export type OfficeSaveMode = 'save' | 'saveAs'\nexport type OfficeAutoSavePolicy = 'disabled' | 'host' | 'editor'",
    "export type OfficeSaveMode = 'save' | 'saveAs'\nexport type OfficeExportFormat = 'docx'\nexport type OfficeAutoSavePolicy = 'disabled' | 'host' | 'editor'",
)
rep(
    path,
    "  saveAs: boolean\n  autoSave: OfficeAutoSavePolicy",
    "  saveAs: boolean\n  saveHistoryVersion: boolean\n  exportDocx: boolean\n  close: boolean\n  autoSave: OfficeAutoSavePolicy",
)
rep(
    path,
    "  saveAs: true,\n  autoSave: 'disabled',",
    "  saveAs: true,\n  saveHistoryVersion: false,\n  exportDocx: true,\n  close: false,\n  autoSave: 'disabled',",
)
rep(
    path,
    "  saveAs: true,\n  autoSave: 'host',",
    "  saveAs: true,\n  saveHistoryVersion: true,\n  exportDocx: true,\n  close: true,\n  autoSave: 'host',",
)
rep(
    path,
    "  baseVersion?: string | null\n  mode?: OfficeSaveMode\n}",
    "  baseVersion?: string | null\n  mode?: OfficeSaveMode\n  /** First persistence of a blank editor document; the Host should choose/create its destination. */\n  newDocument?: boolean\n}",
)
rep(
    path,
    "export interface PickFileOptions {",
    "export interface SaveHistoryVersionInput {\n  file: OfficeFileDescriptor\n  bytes: ArrayBuffer\n  baseVersion?: string | null\n}\n\nexport interface SaveHistoryVersionResult {\n  ok: boolean\n  error?: string\n  code?: 'PERMISSION_DENIED' | 'NOT_FOUND' | 'SAVE_FAILED' | 'CANCELLED'\n}\n\nexport interface ExportDocumentInput {\n  format: OfficeExportFormat\n  file: OfficeFileDescriptor\n  bytes: ArrayBuffer\n}\n\nexport interface ExportDocumentResult {\n  ok: boolean\n  error?: string\n  code?: 'PERMISSION_DENIED' | 'SAVE_FAILED' | 'CANCELLED'\n}\n\nexport interface PickFileOptions {",
)
rep(
    path,
    "  saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult>\n  pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null>",
    "  saveDocument(input: SaveDocumentInput): Promise<SaveDocumentResult>\n  saveHistoryVersion?(input: SaveHistoryVersionInput): Promise<SaveHistoryVersionResult>\n  exportDocument?(input: ExportDocumentInput): Promise<ExportDocumentResult>\n  requestClose?(): Promise<void>\n  pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null>",
)

# 2) Office protocol ---------------------------------------------------------
path = "packages/office-protocol/src/index.ts"
rep(path, "  OfficeEditorMode,\n  OfficeFile,", "  OfficeEditorMode,\n  OfficeExportFormat,\n  OfficeFile,")
rep(
    path,
    "  PickFileOptions,\n  SaveDocumentResult,",
    "  ExportDocumentResult,\n  PickFileOptions,\n  SaveDocumentResult,\n  SaveHistoryVersionResult,",
)
rep(
    path,
    "export interface OfficeEditorState {",
    "export interface OfficeNewPayload {\n  kind: OfficeDocumentKind\n  mode: OfficeEditorMode\n  locale?: string\n  capabilities?: Partial<OfficeHostCapabilities>\n}\n\nexport interface OfficeEditorState {",
)
rep(
    path,
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:set-mode'",
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:new'\n      requestId: string\n      payload: OfficeNewPayload\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:set-locale'\n      requestId?: string\n      payload: { locale: string }\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:set-mode'",
    1,
)
rep(
    path,
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:pick-file-result'",
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:save-history-version-result'\n      requestId: string\n      payload: SaveHistoryVersionResult\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:export-document-result'\n      requestId: string\n      payload: ExportDocumentResult\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:pick-file-result'",
    1,
)
rep(
    path,
    "        baseVersion?: string | null\n        mode?: OfficeSaveMode\n      }",
    "        baseVersion?: string | null\n        mode?: OfficeSaveMode\n        newDocument?: boolean\n      }",
    1,
)
rep(
    path,
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:pick-file'",
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:save-history-version'\n      requestId: string\n      payload: {\n        file: OfficeFileDescriptor\n        bytes: ArrayBuffer\n        baseVersion?: string | null\n      }\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:export-document'\n      requestId: string\n      payload: {\n        format: OfficeExportFormat\n        file: OfficeFileDescriptor\n        bytes: ArrayBuffer\n      }\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:close-request'\n      payload: { reason: 'file-menu' }\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:pick-file'",
    1,
)

# 3) Browser runtimes --------------------------------------------------------
path = "packages/web-runtime/src/index.ts"
rep(
    path,
    "import type {\n  OfficeFile,\n  OfficeHostApi,\n  PickFileOptions,\n  SaveDocumentInput,\n  SaveDocumentResult,\n  SelectedOfficeFile,\n} from '@genoffice/office-host-api'",
    "import type {\n  ExportDocumentInput,\n  ExportDocumentResult,\n  OfficeFile,\n  OfficeHostApi,\n  PickFileOptions,\n  SaveDocumentInput,\n  SaveDocumentResult,\n  SaveHistoryVersionInput,\n  SaveHistoryVersionResult,\n  SelectedOfficeFile,\n} from '@genoffice/office-host-api'",
)
rep(
    path,
    "  async pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null> {",
    "  async saveHistoryVersion(_input: SaveHistoryVersionInput): Promise<SaveHistoryVersionResult> {\n    return {\n      ok: false,\n      code: 'SAVE_FAILED',\n      error: 'History versions require an embedded platform host.',\n    }\n  }\n\n  async exportDocument(input: ExportDocumentInput): Promise<ExportDocumentResult> {\n    downloadBuffer(input.bytes, input.file.name, input.file.mimeType)\n    return { ok: true }\n  }\n\n  async requestClose(): Promise<void> {\n    window.close()\n  }\n\n  async pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null> {",
    1,
)
rep(
    path,
    "          baseVersion: input.baseVersion,\n          mode: input.mode,",
    "          baseVersion: input.baseVersion,\n          mode: input.mode,\n          newDocument: input.newDocument,",
    1,
)
text = read(path)
marker = "  async pickFile(options: PickFileOptions): Promise<SelectedOfficeFile[] | null> {\n    const requestId = createOfficeRequestId('pick')"
idx = text.find(marker, text.find("export class EmbeddedOfficeHost"))
if idx < 0:
    raise SystemExit("embedded pickFile anchor not found")
block = """  async saveHistoryVersion(input: SaveHistoryVersionInput): Promise<SaveHistoryVersionResult> {
    const requestId = createOfficeRequestId('history')
    const bytes = input.bytes.slice(0)
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:save-history-version-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:save-history-version',
        requestId,
        payload: { file: input.file, bytes, baseVersion: input.baseVersion },
      },
      'office:save-history-version-result',
      [bytes],
    )
    return response.payload
  }

  async exportDocument(input: ExportDocumentInput): Promise<ExportDocumentResult> {
    const requestId = createOfficeRequestId('export')
    const bytes = input.bytes.slice(0)
    const response = await this.bridge.request<
      Extract<HostToEditorMessage, { type: 'office:export-document-result' }>
    >(
      {
        protocol: OFFICE_PROTOCOL_VERSION,
        type: 'office:export-document',
        requestId,
        payload: { format: input.format, file: input.file, bytes },
      },
      'office:export-document-result',
      [bytes],
    )
    return response.payload
  }

  async requestClose(): Promise<void> {
    this.bridge.send({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:close-request',
      payload: { reason: 'file-menu' },
    })
  }

"""
write(path, text[:idx] + block + text[idx:])

# 4) Docs DesktopApi and web language surface -------------------------------
path = "apps/docs/src/shared/ipc.ts"
text = read(path)
if "import type { Lang } from '@genoffice/i18n'" not in text:
    text = "import type { Lang } from '@genoffice/i18n'\n\n" + text
old = """  getLanguage(): Promise<'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'>
  /** language switched from the shell home page */
  onLanguageChanged(
    handler: (
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => void,
  ): () => void
"""
new = """  getLanguage(): Promise<Lang>
  /** language switched from the shell or embedded host */
  onLanguageChanged(handler: (lang: Lang) => void): () => void
"""
if old not in text:
    raise SystemExit("DesktopApi language block not found")
text = text.replace(old, new, 1)
old = """  saveDocxNew(
    defaultName: string,
    data: ArrayBuffer,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  getRecentFiles(): Promise<string[]>
"""
new = """  saveDocxNew(
    defaultName: string,
    data: ArrayBuffer,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  /** Platform-owned snapshot; does not replace the current file identity. */
  saveHistoryVersion?(
    defaultName: string,
    data: ArrayBuffer,
  ): Promise<{ ok: boolean; error?: string }>
  /** Export/download current DOCX bytes without changing the current file identity. */
  exportDocx?(defaultName: string, data: ArrayBuffer): Promise<{ ok: boolean; error?: string }>
  /** The embedded platform owns the actual plugin/window lifecycle. */
  requestHostClose?(): Promise<void>
  getRecentFiles(): Promise<string[]>
"""
if old not in text:
    raise SystemExit("DesktopApi saveDocxNew block not found")
write(path, text.replace(old, new, 1))

path = "apps/docs/src/web/desktop-api.ts"
text = read(path)
text = text.replace(
    "import type { ProjectApi } from '@genoffice/project-store'\n",
    "import type { ProjectApi } from '@genoffice/project-store'\nimport { normalizeLang as normalizeUiLang, type Lang } from '@genoffice/i18n'\n",
    1,
)
text = text.replace("type DocsLang = Awaited<ReturnType<DesktopApi['getLanguage']>>", "type DocsLang = Lang", 1)
text = re.sub(
    r"function normalizeLang\(value: string\): DocsLang \{.*?\n\}",
    "function normalizeLang(value: string): DocsLang {\n  return normalizeUiLang(value)\n}",
    text,
    count=1,
    flags=re.S,
)
anchor = """  const setMode = (nextMode: 'view' | 'edit'): void => {
    if (mode === nextMode) return
    mode = nextMode
    for (const handler of modeHandlers) handler(mode)
  }

"""
if anchor not in text:
    raise SystemExit("setMode block not found")
text = text.replace(
    anchor,
    anchor
    + """  const setLanguage = (locale: string): void => {
    const next = normalizeLang(locale)
    if (currentLang === next) return
    currentLang = next
    for (const handler of languageHandlers) handler(currentLang)
  }

""",
    1,
)
old = """    saving = true
    try {
      const result = await host.saveDocument({
        file,
        bytes: data,
        baseVersion: current?.file.version ?? null,
      })
"""
new = """    const newDocument = current === null
    saving = true
    try {
      const result = await host.saveDocument({
        file,
        bytes: data,
        baseVersion: current?.file.version ?? null,
        newDocument,
      })
"""
if old not in text:
    raise SystemExit("saveWithName host call not found")
text = text.replace(old, new, 1)
old = "    saveDocxNew: async (defaultName, data) => saveWithName(defaultName, data),\n    getRecentFiles: async () => [],"
new = """    saveDocxNew: async (defaultName, data) => saveWithName(defaultName, data),
    saveHistoryVersion: async (_defaultName, data) => {
      if (!current) return { ok: false, error: 'Save the new document before creating history.' }
      if (!host.saveHistoryVersion) {
        return { ok: false, error: 'History versions are not supported by this host.' }
      }
      const result = await host.saveHistoryVersion({
        file: current.file,
        bytes: data,
        baseVersion: current.file.version ?? null,
      })
      return { ok: result.ok, error: result.error }
    },
    exportDocx: async (defaultName, data) => {
      if (!host.exportDocument) {
        return { ok: false, error: 'DOCX export is not supported by this host.' }
      }
      const descriptor: OfficeFileDescriptor = current?.file ?? {
        id: `export:${Date.now()}`,
        name: defaultName,
        mimeType: DOCX_MIME,
        size: data.byteLength,
        version: null,
      }
      const result = await host.exportDocument({
        format: 'docx',
        file: { ...descriptor, name: defaultName, size: data.byteLength },
        bytes: data,
      })
      return { ok: result.ok, error: result.error }
    },
    requestHostClose: async () => {
      await host.requestClose?.()
    },
    getRecentFiles: async () => [],"""
if old not in text:
    raise SystemExit("web desktop save-new block not found")
text = text.replace(old, new, 1)
old = """      case 'office:init': {
        if (message.payload.kind !== 'docx') return
        setMode(message.payload.mode)
        if (message.payload.locale) {
          currentLang = normalizeLang(message.payload.locale)
          for (const handler of languageHandlers) handler(currentLang)
        }
        void setCurrentFile(message.payload.file).then((result) => {
"""
new = """      case 'office:init': {
        if (message.payload.kind !== 'docx') return
        setMode(message.payload.mode)
        if (message.payload.locale) setLanguage(message.payload.locale)
        void setCurrentFile(message.payload.file).then((result) => {
"""
if old not in text:
    raise SystemExit("office:init locale block not found")
text = text.replace(old, new, 1)
anchor = "      case 'office:save': {\n"
idx = text.find(anchor)
if idx < 0:
    raise SystemExit("office:save switch anchor not found")
addition = """      case 'office:new': {
        if (message.payload.kind !== 'docx') return
        setMode(message.payload.mode)
        if (message.payload.locale) setLanguage(message.payload.locale)
        current = null
        pendingOpen = null
        if (initialOpenResolve) {
          const resolve = initialOpenResolve
          initialOpenResolve = null
          resolve(null)
        } else {
          for (const handler of menuHandlers) handler('new' satisfies MenuCommand)
        }
        break
      }
      case 'office:set-locale':
        setLanguage(message.payload.locale)
        break
"""
text = text[:idx] + addition + text[idx:]
write(path, text)

# 5) Capability policy -------------------------------------------------------
path = "apps/docs/src/web/host-policy.ts"
rep(
    path,
    "  'office-can-save-as',\n] as const",
    "  'office-can-save-as',\n  'office-can-save-history',\n  'office-can-export-docx',\n  'office-can-close',\n] as const",
)
rep(
    path,
    "  root.classList.toggle('office-can-save-as', capabilities.saveAs)\n}",
    "  root.classList.toggle('office-can-save-as', capabilities.saveAs)\n  root.classList.toggle('office-can-save-history', capabilities.saveHistoryVersion)\n  root.classList.toggle('office-can-export-docx', capabilities.exportDocx)\n  root.classList.toggle('office-can-close', capabilities.close)\n}",
)
rep(
    path,
    "    if (message.type === 'office:init' && message.payload.capabilities) {\n      apply(message.payload.capabilities)\n    }",
    "    if (\n      (message.type === 'office:init' || message.type === 'office:new') &&\n      message.payload.capabilities\n    ) {\n      apply(message.payload.capabilities)\n    }",
)

path = "apps/docs/src/web/product-policy.css"
rep(
    path,
    "html.office-web:not(.office-can-open) .file-menu > button:nth-child(1),\nhtml.office-web:not(.office-can-save) .file-menu > button:nth-child(2),\nhtml.office-web:not(.office-can-save-as) .file-menu > button:nth-child(3) {\n  display: none;\n}",
    "html.office-web:not(.office-can-open) .file-menu-open,\nhtml.office-web:not(.office-can-save) .file-menu-save,\nhtml.office-web:not(.office-can-save-as) .file-menu-save-as,\nhtml.office-web:not(.office-can-save-history) .file-menu-save-history,\nhtml.office-web:not(.office-can-export-docx) .file-menu-export-docx,\nhtml.office-web:not(.office-can-close) .file-menu-exit {\n  display: none;\n}",
)

# 6) Ribbon File menu --------------------------------------------------------
path = "apps/docs/src/renderer/components/Ribbon.tsx"
rep(
    path,
    "  onOpen: () => void\n  onSave: () => void\n  onSaveAs: () => void\n  showAi: boolean",
    "  onOpen: () => void\n  onSave: () => void\n  onSaveAs: () => void\n  onSaveHistoryVersion?: () => void\n  onExportDocx?: () => void\n  onExit?: () => void\n  showAi: boolean",
)
rep(
    path,
    "  onOpen,\n  onSave,\n  onSaveAs,\n  showAi,",
    "  onOpen,\n  onSave,\n  onSaveAs,\n  onSaveHistoryVersion,\n  onExportDocx,\n  onExit,\n  showAi,",
)
text = read(path)
start_marker = "            {dropdown === 'file' && (\n"
end_marker = "            )}\n          </div>"
start = text.find(start_marker)
if start < 0:
    raise SystemExit("File menu start not found")
end = text.find(end_marker, start)
if end < 0:
    raise SystemExit("File menu end not found")
end += len("            )}\n")
menu = """            {dropdown === 'file' && (
              <div className="file-menu">
                <button
                  className="file-menu-open"
                  onClick={() => {
                    setDropdown(null)
                    onOpen()
                  }}
                >
                  {t('ribbonOpen')} <span className="file-menu-key">Ctrl+O</span>
                </button>
                <button
                  className="file-menu-save"
                  disabled={!hasDoc}
                  onClick={() => {
                    setDropdown(null)
                    onSave()
                  }}
                >
                  {t('ribbonSave')} <span className="file-menu-key">Ctrl+S</span>
                </button>
                <button
                  className="file-menu-save-as"
                  disabled={!hasDoc}
                  onClick={() => {
                    setDropdown(null)
                    onSaveAs()
                  }}
                >
                  {t('ribbonSaveAs')} <span className="file-menu-key">Ctrl+Shift+S</span>
                </button>
                {onSaveHistoryVersion && (
                  <button
                    className="file-menu-save-history"
                    disabled={!hasDoc || !filePath}
                    onClick={() => {
                      setDropdown(null)
                      onSaveHistoryVersion()
                    }}
                  >
                    {t('ribbonSaveHistoryVersion')}
                  </button>
                )}
                {onExportDocx && (
                  <button
                    className="file-menu-export-docx"
                    disabled={!hasDoc}
                    onClick={() => {
                      setDropdown(null)
                      onExportDocx()
                    }}
                  >
                    {t('ribbonExportDocx')}
                  </button>
                )}
                {onExit && (
                  <button
                    className="file-menu-exit"
                    onClick={() => {
                      setDropdown(null)
                      onExit()
                    }}
                  >
                    {t('ribbonExit')}
                  </button>
                )}
              </div>
            )}
"""
write(path, text[:start] + menu + text[end:])

# 7) i18n --------------------------------------------------------------------
path = "apps/docs/src/renderer/i18n/strings-ribbon.ts"
lines = read(path).splitlines()
translations = {
    "zh": ["存为历史版本", "导出为 DOCX", "退出", "退出前保存更改？", "此文档有尚未保存的更改。", "保存并退出", "放弃并退出"],
    "en": ["Save as History Version", "Export as DOCX", "Exit", "Save changes before exiting?", "This document has unsaved changes.", "Save and Exit", "Discard and Exit"],
    "ja": ["履歴バージョンとして保存", "DOCX としてエクスポート", "終了", "終了前に変更を保存しますか？", "この文書には未保存の変更があります。", "保存して終了", "破棄して終了"],
    "ko": ["기록 버전으로 저장", "DOCX로 내보내기", "종료", "종료하기 전에 변경 사항을 저장하시겠습니까?", "이 문서에 저장되지 않은 변경 사항이 있습니다.", "저장 후 종료", "버리고 종료"],
    "fr": ["Enregistrer comme version d’historique", "Exporter en DOCX", "Quitter", "Enregistrer les modifications avant de quitter ?", "Ce document contient des modifications non enregistrées.", "Enregistrer et quitter", "Ignorer et quitter"],
    "de": ["Als Verlaufsversion speichern", "Als DOCX exportieren", "Beenden", "Änderungen vor dem Beenden speichern?", "Dieses Dokument enthält nicht gespeicherte Änderungen.", "Speichern und beenden", "Verwerfen und beenden"],
    "es": ["Guardar como versión del historial", "Exportar como DOCX", "Salir", "¿Guardar los cambios antes de salir?", "Este documento tiene cambios sin guardar.", "Guardar y salir", "Descartar y salir"],
    "th": ["บันทึกเป็นเวอร์ชันประวัติ", "ส่งออกเป็น DOCX", "ออก", "บันทึกการเปลี่ยนแปลงก่อนออกหรือไม่", "เอกสารนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก", "บันทึกและออก", "ละทิ้งและออก"],
    "id": ["Simpan sebagai Versi Riwayat", "Ekspor sebagai DOCX", "Keluar", "Simpan perubahan sebelum keluar?", "Dokumen ini memiliki perubahan yang belum disimpan.", "Simpan dan Keluar", "Buang dan Keluar"],
    "ru": ["Сохранить как версию истории", "Экспортировать в DOCX", "Выйти", "Сохранить изменения перед выходом?", "В документе есть несохранённые изменения.", "Сохранить и выйти", "Не сохранять и выйти"],
    "ar": ["حفظ كإصدار في السجل", "تصدير بصيغة DOCX", "خروج", "هل تريد حفظ التغييرات قبل الخروج؟", "يحتوي هذا المستند على تغييرات غير محفوظة.", "حفظ وخروج", "تجاهل وخروج"],
    "pt": ["Salvar como versão do histórico", "Exportar como DOCX", "Sair", "Salvar alterações antes de sair?", "Este documento tem alterações não salvas.", "Salvar e sair", "Descartar e sair"],
    "it": ["Salva come versione cronologia", "Esporta come DOCX", "Esci", "Salvare le modifiche prima di uscire?", "Questo documento contiene modifiche non salvate.", "Salva ed esci", "Ignora ed esci"],
    "pl": ["Zapisz jako wersję historii", "Eksportuj jako DOCX", "Wyjdź", "Zapisać zmiany przed wyjściem?", "Ten dokument zawiera niezapisane zmiany.", "Zapisz i wyjdź", "Odrzuć i wyjdź"],
    "nl": ["Opslaan als geschiedenisversie", "Exporteren als DOCX", "Afsluiten", "Wijzigingen opslaan voordat u afsluit?", "Dit document bevat niet-opgeslagen wijzigingen.", "Opslaan en afsluiten", "Negeren en afsluiten"],
    "ms": ["Simpan sebagai Versi Sejarah", "Eksport sebagai DOCX", "Keluar", "Simpan perubahan sebelum keluar?", "Dokumen ini mempunyai perubahan yang belum disimpan.", "Simpan dan Keluar", "Buang dan Keluar"],
    "he": ["שמירה כגרסת היסטוריה", "ייצוא כ-DOCX", "יציאה", "לשמור שינויים לפני היציאה?", "במסמך זה יש שינויים שלא נשמרו.", "שמירה ויציאה", "ביטול שינויים ויציאה"],
    "hi": ["इतिहास संस्करण के रूप में सहेजें", "DOCX के रूप में निर्यात करें", "बाहर निकलें", "बाहर निकलने से पहले बदलाव सहेजें?", "इस दस्तावेज़ में सहेजे न गए बदलाव हैं।", "सहेजें और बाहर निकलें", "छोड़ें और बाहर निकलें"],
    "zh-TW": ["存為歷史版本", "匯出為 DOCX", "退出", "退出前儲存變更？", "此文件有尚未儲存的變更。", "儲存並退出", "放棄並退出"],
}
keys = [
    "ribbonSaveHistoryVersion",
    "ribbonExportDocx",
    "ribbonExit",
    "ribbonExitUnsavedTitle",
    "ribbonExitUnsavedMessage",
    "ribbonExitSave",
    "ribbonExitDiscard",
]
current = None
inserted = set()
out = []
block_re = re.compile(r"^  (?:(?:'([^']+)')|([A-Za-z-]+)): \{$")
for line in lines:
    m = block_re.match(line)
    if m:
        current = m.group(1) or m.group(2)
    out.append(line)
    if current in translations and "ribbonSaveAs:" in line:
        indent = line[: len(line) - len(line.lstrip())]
        for key, value in zip(keys, translations[current]):
            escaped = value.replace("\\", "\\\\").replace("'", "\\'")
            out.append(f"{indent}{key}: '{escaped}',")
        inserted.add(current)
missing = set(translations) - inserted
if missing:
    raise SystemExit(f"i18n anchors missing: {sorted(missing)}")
write(path, "\n".join(out) + "\n")

# 8) Exit confirmation UI ---------------------------------------------------
write(
    "apps/docs/src/renderer/components/ExitConfirmModal.tsx",
    """import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

export function ExitConfirmModal({
  saving,
  onSave,
  onDiscard,
  onCancel,
}: {
  saving: boolean
  onSave: () => void
  onDiscard: () => void
  onCancel: () => void
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onCancel)
  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(event) => event.target === event.currentTarget && !saving && onCancel()}
    >
      <div className="modal">
        <h2>{t('ribbonExitUnsavedTitle')}</h2>
        <p>{t('ribbonExitUnsavedMessage')}</p>
        <div className="modal-actions">
          <button className="btn-ghost" disabled={saving} onClick={onCancel}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-ghost" disabled={saving} onClick={onDiscard}>
            {t('ribbonExitDiscard')}
          </button>
          <button className="btn-primary" disabled={saving} onClick={onSave}>
            {t('ribbonExitSave')}
          </button>
        </div>
      </div>
    </div>
  )
}
""",
)

# 9) App actions -------------------------------------------------------------
path = "apps/docs/src/renderer/App.tsx"
text = read(path)
text = text.replace(
    "import { PromptModal } from './components/PromptModal'\n",
    "import { PromptModal } from './components/PromptModal'\nimport { ExitConfirmModal } from './components/ExitConfirmModal'\n",
    1,
)
text = text.replace(
    "import {\n  exportPdf as exportPdfImpl,\n",
    "import {\n  buildDocBytes,\n  exportPdf as exportPdfImpl,\n",
    1,
)
save_anchor = """  const save = useCallback(
    (saveAs: boolean, auto = false) => saveImpl(fileCtxRef.current, saveAs, auto),
    [],
  )

"""
if save_anchor not in text:
    raise SystemExit("App save callback anchor not found")
actions = """  const save = useCallback(
    (saveAs: boolean, auto = false) => saveImpl(fileCtxRef.current, saveAs, auto),
    [],
  )

  const [showExitConfirm, setShowExitConfirm] = useState(false)
  const [exitSaving, setExitSaving] = useState(false)

  const currentDocBuffer = useCallback(async (): Promise<ArrayBuffer | null> => {
    const bytes = await buildDocBytes(fileCtxRef.current)
    if (!bytes) return null
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  }, [])

  const saveHistoryVersion = useCallback(async () => {
    const current = fileCtxRef.current.doc
    if (!current?.filePath || !window.desktop.saveHistoryVersion) return
    const buffer = await currentDocBuffer()
    if (!buffer) return
    await window.desktop.saveHistoryVersion(current.fileName, buffer)
  }, [currentDocBuffer])

  const exportDocx = useCallback(async () => {
    const current = fileCtxRef.current.doc
    if (!current || !window.desktop.exportDocx) return
    const buffer = await currentDocBuffer()
    if (!buffer) return
    await window.desktop.exportDocx(current.fileName, buffer)
  }, [currentDocBuffer])

  const requestExit = useCallback(() => {
    if (!window.desktop.requestHostClose) return
    if (isDocDirty(fileCtxRef.current)) {
      setShowExitConfirm(true)
      return
    }
    void window.desktop.requestHostClose()
  }, [])

  const discardAndExit = useCallback(() => {
    setShowExitConfirm(false)
    void window.desktop.requestHostClose?.()
  }, [])

  const saveAndExit = useCallback(async () => {
    if (exitSaving) return
    setExitSaving(true)
    try {
      const ok = await save(false)
      if (ok && !isDocDirty(fileCtxRef.current)) {
        setShowExitConfirm(false)
        await window.desktop.requestHostClose?.()
      }
    } finally {
      setExitSaving(false)
    }
  }, [exitSaving, save])

"""
text = text.replace(save_anchor, actions, 1)
ribbon_anchor = """          showGrid={showGrid}
          splitView={splitView}
          {...ribbonActions}
"""
if ribbon_anchor not in text:
    raise SystemExit("App Ribbon props anchor not found")
text = text.replace(
    ribbon_anchor,
    """          showGrid={showGrid}
          splitView={splitView}
          onSaveHistoryVersion={
            window.desktop.saveHistoryVersion ? () => void saveHistoryVersion() : undefined
          }
          onExportDocx={window.desktop.exportDocx ? () => void exportDocx() : undefined}
          onExit={window.desktop.requestHostClose ? requestExit : undefined}
          {...ribbonActions}
""",
    1,
)
modal_anchor = "      {showLinkModal && <LinkInsertModal editor={editor} onClose={() => setShowLinkModal(false)} />}\n"
if modal_anchor not in text:
    raise SystemExit("App modal anchor not found")
text = text.replace(
    modal_anchor,
    """      {showExitConfirm && (
        <ExitConfirmModal
          saving={exitSaving}
          onSave={() => void saveAndExit()}
          onDiscard={discardAndExit}
          onCancel={() => setShowExitConfirm(false)}
        />
      )}

"""
    + modal_anchor,
    1,
)
write(path, text)

# 10) Unit regression --------------------------------------------------------
path = "apps/docs/tests/web-desktop-api.test.ts"
text = read(path)
anchor = "  it('routes parent save through the host and acknowledges the original request', async () => {"
if anchor not in text:
    raise SystemExit("web adapter test anchor not found")
test = """  it('starts a blank embedded document, switches locale live, and marks its first save', async () => {
    const { controller, emit, saveDocument, destroy } = createHarness()
    const changed = vi.fn()
    const offLanguage = controller.desktopApi.onLanguageChanged(changed)
    const pendingOpen = controller.desktopApi.consumePendingOpenDocx()

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:new',
      requestId: 'new-1',
      payload: { kind: 'docx', mode: 'edit', locale: 'pt-BR', capabilities: { open: true } },
    })

    expect(await pendingOpen).toBeNull()
    expect(await controller.desktopApi.getLanguage()).toBe('pt')
    expect(changed).toHaveBeenCalledWith('pt')

    emit({
      protocol: OFFICE_PROTOCOL_VERSION,
      type: 'office:set-locale',
      payload: { locale: 'zh-TW' },
    })
    expect(await controller.desktopApi.getLanguage()).toBe('zh-TW')
    expect(changed).toHaveBeenLastCalledWith('zh-TW')

    await controller.desktopApi.saveDocxNew('新建文档.docx', bytesOf('draft'))
    expect(saveDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        newDocument: true,
        file: expect.objectContaining({ name: '新建文档.docx' }),
      }),
    )

    offLanguage()
    destroy()
  })

"""
write(path, text.replace(anchor, test + anchor, 1))
