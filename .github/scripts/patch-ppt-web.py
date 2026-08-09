from pathlib import Path
import json
import re

ROOT = Path(__file__).resolve().parents[2]

def read(path): return (ROOT / path).read_text(encoding='utf-8')
def write(path, text): (ROOT / path).write_text(text, encoding='utf-8')

def must_replace(text, old, new, label):
    if old not in text:
        raise RuntimeError(f'missing anchor: {label}')
    return text.replace(old, new, 1)

# Root scripts.
p = ROOT / 'package.json'
data = json.loads(p.read_text())
s = data['scripts']
s['dev:web:slides'] = 'npm run dev:web -w @genoffice/slides'
s['build:web:slides'] = 'npm run build:web -w @genoffice/slides'
p.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n')

# Slides API browser hooks + shared language type.
path = 'apps/slides/src/shared/ipc.ts'
text = read(path)
if "import type { Lang } from '@genoffice/i18n'" not in text:
    text = text.replace("import type { RenderSlide } from '@genoffice/pptx-render'\n", "import type { RenderSlide } from '@genoffice/pptx-render'\nimport type { Lang } from '@genoffice/i18n'\n", 1)
text = re.sub(
    r"  getLanguage: \(\) => Promise<\n    'zh' \| 'en' \| 'ja' \| 'ko' \| 'fr' \| 'de' \| 'es' \| 'th' \| 'id' \| 'ru' \| 'ar'\n  >\n  /\*\* language switched from the shell home page \*/\n  onLanguageChanged: \(\n    handler: \(\n      lang: 'zh' \| 'en' \| 'ja' \| 'ko' \| 'fr' \| 'de' \| 'es' \| 'th' \| 'id' \| 'ru' \| 'ar',\n    \) => void,\n  \) => \(\) => void\n",
    "  getLanguage: () => Promise<Lang>\n  /** language switched from the shell home page */\n  onLanguageChanged: (handler: (lang: Lang) => void) => () => void\n  /** Optional Web Host-enforced editor mode. */\n  getHostEditorMode?: () => Promise<'view' | 'edit'>\n  onHostEditorModeChanged?: (handler: (mode: 'view' | 'edit') => void) => () => void\n  /** Web Host dirty-state mirror. */\n  reportDirtyChange?: (dirty: boolean) => void\n",
    text,
    count=1,
)
text = must_replace(
    text,
    "  saveAs: (\n    defaultName: string,\n  ) => Promise<{ ok: boolean; path?: string; error?: string; slides?: RenderSlide[] }>\n",
    "  saveAs: (\n    defaultName: string,\n  ) => Promise<{ ok: boolean; path?: string; error?: string; slides?: RenderSlide[] }>\n  /** Web Host lifecycle extensions; absent in the legacy Electron preload. */\n  saveHistoryVersion?: () => Promise<{ ok: boolean; error?: string }>\n  exportPptx?: () => Promise<{ ok: boolean; error?: string }>\n  requestHostClose?: () => Promise<void>\n",
    'slides api lifecycle hooks',
)
write(path, text)

# Fix current Web controller type errors from the first diagnostic.
path = 'apps/slides/src/web/slides-api.ts'
text = read(path)
text = text.replace(
    "      const el = slide?.elements.find((item) => item.id === op.sourceId)\n      if (!slide || !el) return null\n      await pushHistory()\n      if (typeof op.fill === 'string') {",
    "      const el = slide?.elements.find((item) => item.id === op.sourceId) as any\n      if (!slide || !el) return null\n      await pushHistory()\n      if (typeof op.fill === 'string') {",
    1,
)
# second occurrence is stroke
anchor = "      const el = slide?.elements.find((item) => item.id === op.sourceId)\n      if (!slide || !el) return null\n      await pushHistory()\n      el.stroke = op.stroke"
if anchor in text:
    text = text.replace(anchor, anchor.replace("const el = slide?.elements.find((item) => item.id === op.sourceId)", "const el = slide?.elements.find((item) => item.id === op.sourceId) as any"), 1)
text = text.replace("  document.documentElement.classList.toggle('office-view-mode', mode === 'view')\n", "  document.documentElement.classList.remove('office-view-mode')\n", 1)
write(path, text)

# Browser-safe Buffer text/base64 conversions across pptx-engine.
def transform_buffers(text):
    # compound conversions first
    text = re.sub(r"Buffer\.from\((.*?),\s*['\"]base64['\"]\)\.toString\(['\"]utf8['\"]\)", r"decodeUtf8(decodeBase64(\1))", text)
    text = re.sub(r"Buffer\.from\((.*?),\s*['\"]utf8['\"]\)\.toString\(['\"]base64['\"]\)", r"encodeBase64(encodeUtf8(\1))", text)
    text = re.sub(r"Buffer\.from\(([^\n]+?)\)\.toString\(['\"]base64['\"]\)", r"encodeBase64(new Uint8Array(\1))", text)
    text = re.sub(r"Buffer\.from\((.*?),\s*['\"]utf8['\"]\)", r"encodeUtf8(\1)", text, flags=re.S)
    text = re.sub(r"Buffer\.from\((.*?),\s*['\"]base64['\"]\)", r"decodeBase64(\1)", text, flags=re.S)
    text = re.sub(r"Buffer\.from\((.*?),\s*['\"]ascii['\"]\)", r"encodeAscii(\1)", text, flags=re.S)
    return text

def add_byte_import(text):
    helpers = [h for h in ['encodeUtf8','decodeUtf8','encodeBase64','decodeBase64','encodeAscii'] if re.search(rf'\b{h}\(', text)]
    # Don't self-import in bytes.ts
    if not helpers or "from './bytes'" in text:
        return text
    imp = "import { " + ', '.join(helpers) + " } from './bytes'\n"
    # Insert after leading comment, before first existing import.
    idx = text.find('import ')
    if idx < 0: return imp + text
    return text[:idx] + imp + text[idx:]

for file in (ROOT / 'packages/pptx-engine/src').glob('*.ts'):
    if file.name in {'bytes.ts', 'media-insert.ts'}:
        continue
    src = file.read_text(encoding='utf-8')
    out = transform_buffers(src)
    if out != src:
        out = add_byte_import(out)
        file.write_text(out, encoding='utf-8')

# Sections: Web Crypto UUID instead of node:crypto.
path = ROOT / 'packages/pptx-engine/src/sections.ts'
text = path.read_text(encoding='utf-8')
text = text.replace("import { randomUUID } from 'node:crypto'\n", '')
text = re.sub(r'\brandomUUID\(\)', 'globalThis.crypto.randomUUID()', text)
path.write_text(text, encoding='utf-8')

# media-insert: remove node:zlib + Buffer-only PNG implementation.
path = ROOT / 'packages/pptx-engine/src/media-insert.ts'
text = path.read_text(encoding='utf-8')
text = text.replace("import { deflateSync } from 'node:zlib'\n", '')
if "from './bytes'" not in text:
    idx = text.find('import ')
    text = text[:idx] + "import { concatBytes, encodeAscii, encodeUtf8, writeUint32Be } from './bytes'\n" + text[idx:]
start = text.index('function pngChunk(')
end = text.index('// ── Shared part / rels surgery', start)
replacement = r'''function adler32(data: Uint8Array): number {
  let a = 1
  let b = 0
  for (const value of data) {
    a = (a + value) % 65521
    b = (b + a) % 65521
  }
  return ((b << 16) | a) >>> 0
}

/** Minimal zlib stream using uncompressed DEFLATE blocks (browser + Node, sync). */
function deflateStored(data: Uint8Array): Uint8Array {
  const parts: Uint8Array[] = [new Uint8Array([0x78, 0x01])]
  for (let offset = 0; offset < data.length; offset += 65535) {
    const block = data.subarray(offset, Math.min(offset + 65535, data.length))
    const final = offset + block.length >= data.length
    const header = new Uint8Array(5)
    header[0] = final ? 1 : 0
    const len = block.length
    header[1] = len & 0xff
    header[2] = (len >>> 8) & 0xff
    const nlen = (~len) & 0xffff
    header[3] = nlen & 0xff
    header[4] = (nlen >>> 8) & 0xff
    parts.push(header, block)
  }
  const checksum = new Uint8Array(4)
  writeUint32Be(checksum, 0, adler32(data))
  parts.push(checksum)
  return concatBytes(parts)
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = encodeAscii(type)
  const head = new Uint8Array(8)
  writeUint32Be(head, 0, data.length)
  head.set(typeBytes, 4)
  const crcBuf = new Uint8Array(4)
  writeUint32Be(crcBuf, 0, crc32(concatBytes([typeBytes, data])))
  return concatBytes([head, data, crcBuf])
}

/** Generate a w×h solid-color PNG (RGB, no alpha). Used as a poster frame placeholder. */
export function solidPng(w: number, h: number, rgb: [number, number, number]): Uint8Array {
  const ihdr = new Uint8Array(13)
  writeUint32Be(ihdr, 0, w)
  writeUint32Be(ihdr, 4, h)
  ihdr[8] = 8
  ihdr[9] = 2
  const raw = new Uint8Array(h * (1 + w * 3))
  for (let y = 0; y < h; y++) {
    const row = y * (1 + w * 3)
    raw[row] = 0
    for (let x = 0; x < w; x++) {
      raw[row + 1 + x * 3] = rgb[0]
      raw[row + 2 + x * 3] = rgb[1]
      raw[row + 3 + x * 3] = rgb[2]
    }
  }
  return concatBytes([
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', deflateStored(raw)),
    pngChunk('IEND', new Uint8Array()),
  ])
}

'''
text = text[:start] + replacement + text[end:]
text = transform_buffers(text)
# transform may require encode helpers already covered by explicit import
path.write_text(text, encoding='utf-8')

# Ribbon lifecycle Props.
path = 'apps/slides/src/renderer/components/ribbon-shared.tsx'
text = read(path)
text = must_replace(
    text,
    "  onSaveAs: () => void\n  /** Export as PDF (hidden slides skipped) */\n",
    "  onSaveAs: () => void\n  /** Web Host lifecycle extensions; omitted by the Electron renderer. */\n  onSaveHistoryVersion?: () => void\n  onExportPptx?: () => void\n  onExit?: () => void\n  /** Export as PDF (hidden slides skipped) */\n",
    'ribbon props web lifecycle',
)
write(path, text)

# Ribbon File menu and labels.
path = 'apps/slides/src/renderer/components/Ribbon.tsx'
text = read(path)
if "slidesWebLifecycleLabels" not in text:
    text = text.replace("import { useI18n, type StringKey } from '../i18n/locale'\n", "import { useI18n, type StringKey } from '../i18n/locale'\nimport { slidesWebLifecycleLabels } from '../web-labels'\n", 1)
text = text.replace("  onSaveAs,\n  onExportPdf,", "  onSaveAs,\n  onSaveHistoryVersion,\n  onExportPptx,\n  onExit,\n  onExportPdf,", 1)
text = text.replace("  const { t } = useI18n()\n", "  const { t, lang } = useI18n()\n  const webLabels = slidesWebLifecycleLabels(lang)\n", 1)
old_menu = '''              <div className="file-menu">
                <button
                  onClick={() => {
                    setFileOpen(false)
                    onOpen()
                  }}
                >
                  {t('ribbonFileOpen')} <span className="file-menu-key">Ctrl+O</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onSave()
                  }}
                >
                  {t('ribbonFileSave')} <span className="file-menu-key">Ctrl+S</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onSaveAs()
                  }}
                >
                  {t('ribbonFileSaveAs')} <span className="file-menu-key">Ctrl+Shift+S</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onExportPdf()
                  }}
                >
                  {t('ribbonFileExportPdf')}
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onPrint()
                  }}
                >
                  {t('ribbonFilePrint')} <span className="file-menu-key">Ctrl+P</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onExportImages()
                  }}
                >
                  {t('ribbonFileExportImages')}
                </button>
              </div>'''
new_menu = '''              <div className="file-menu">
                <button
                  className="file-menu-open"
                  onClick={() => {
                    setFileOpen(false)
                    onOpen()
                  }}
                >
                  {t('ribbonFileOpen')} <span className="file-menu-key">Ctrl+O</span>
                </button>
                <button
                  className="file-menu-save"
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onSave()
                  }}
                >
                  {t('ribbonFileSave')} <span className="file-menu-key">Ctrl+S</span>
                </button>
                <button
                  className="file-menu-save-as"
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onSaveAs()
                  }}
                >
                  {t('ribbonFileSaveAs')} <span className="file-menu-key">Ctrl+Shift+S</span>
                </button>
                {onSaveHistoryVersion && (
                  <button
                    className="file-menu-save-history"
                    disabled={!hasDoc}
                    onClick={() => {
                      setFileOpen(false)
                      onSaveHistoryVersion()
                    }}
                  >
                    {webLabels.saveHistory}
                  </button>
                )}
                {onExportPptx && (
                  <button
                    className="file-menu-export-pptx"
                    disabled={!hasDoc}
                    onClick={() => {
                      setFileOpen(false)
                      onExportPptx()
                    }}
                  >
                    {webLabels.exportPptx}
                  </button>
                )}
                {onExit && (
                  <button
                    className="file-menu-exit"
                    onClick={() => {
                      setFileOpen(false)
                      onExit()
                    }}
                  >
                    {webLabels.exit}
                  </button>
                )}
                <button
                  className="file-menu-electron-only"
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onExportPdf()
                  }}
                >
                  {t('ribbonFileExportPdf')}
                </button>
                <button
                  className="file-menu-electron-only"
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onPrint()
                  }}
                >
                  {t('ribbonFilePrint')} <span className="file-menu-key">Ctrl+P</span>
                </button>
                <button
                  className="file-menu-electron-only"
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onExportImages()
                  }}
                >
                  {t('ribbonFileExportImages')}
                </button>
              </div>'''
text = must_replace(text, old_menu, new_menu, 'slides file menu')
write(path, text)

# App: host mode, dirty mirror, lifecycle callbacks + exit dialog.
path = 'apps/slides/src/renderer/App.tsx'
text = read(path)
if "slidesWebLifecycleLabels" not in text:
    # Place alongside file-actions import if present; otherwise before component imports.
    marker = "import * as fileActions from './file-actions'\n"
    if marker in text:
        text = text.replace(marker, marker + "import { slidesWebLifecycleLabels } from './web-labels'\n", 1)
    else:
        idx = text.find("export function App()")
        text = text[:idx] + "import { slidesWebLifecycleLabels } from './web-labels'\n" + text[idx:]
# useI18n currently provides lang near App beginning
text = text.replace("  const { t } = useI18n()\n", "  const { t, lang } = useI18n()\n  const webLabels = slidesWebLifecycleLabels(lang)\n", 1)
text = must_replace(
    text,
    "  const [viewMode, setViewMode] = useState<SlidesViewMode>('normal')\n",
    "  const [viewMode, setViewMode] = useState<SlidesViewMode>('normal')\n  const [hostEditorMode, setHostEditorMode] = useState<'view' | 'edit'>('edit')\n  const hostForcedReadingRef = useRef(false)\n  const [exitConfirmOpen, setExitConfirmOpen] = useState(false)\n  const [exitSaving, setExitSaving] = useState(false)\n",
    'view mode state',
)
# Add mode effects after applyOpen block marker.
marker = "  const setSelectedId = useCallback((id: string | null, additive = false) => {\n"
effect = '''  // Web Host mode is authoritative. View mode maps to the existing reading view,
  // so the editing canvas and editing Ribbon are not reachable while the Host is read-only.
  useEffect(() => {
    let active = true
    void window.slidesApi.getHostEditorMode?.().then((next) => {
      if (active && next) setHostEditorMode(next)
    })
    const off = window.slidesApi.onHostEditorModeChanged?.((next) => setHostEditorMode(next))
    return () => {
      active = false
      off?.()
    }
  }, [])

  useEffect(() => {
    if (hostEditorMode === 'view') {
      hostForcedReadingRef.current = true
      if (viewMode !== 'reading') setViewMode('reading')
      return
    }
    if (hostForcedReadingRef.current) {
      hostForcedReadingRef.current = false
      if (viewMode === 'reading') setViewMode('normal')
    }
  }, [hostEditorMode, viewMode])

  useEffect(() => {
    window.slidesApi.reportDirtyChange?.(dirty)
  }, [dirty])

'''
text = must_replace(text, marker, effect + marker, 'host mode effects')
# Lifecycle callbacks after exportPdf definitions.
marker = "  const exportPdf = useCallback(() => fileActions.exportPdf(ctxRef.current), [])\n"
lifecycle = '''  const saveHistoryVersion = useCallback(async () => {
    await fileActions.flushActiveEdit(ctxRef.current)
    await ctxRef.current.flushNotes()
    const result = await window.slidesApi.saveHistoryVersion?.()
    if (!result) return
    setStatus(result.ok ? webLabels.historySaved : `${webLabels.historyFailed}: ${result.error ?? ''}`)
  }, [webLabels])

  const exportPptx = useCallback(async () => {
    await fileActions.flushActiveEdit(ctxRef.current)
    await ctxRef.current.flushNotes()
    const result = await window.slidesApi.exportPptx?.()
    if (!result) return
    setStatus(result.ok ? webLabels.exportDone : `${webLabels.exportFailed}: ${result.error ?? ''}`)
  }, [webLabels])

  const requestExit = useCallback(async () => {
    if (!dirty) {
      await window.slidesApi.requestHostClose?.()
      return
    }
    setExitConfirmOpen(true)
  }, [dirty])

  const saveAndExit = useCallback(async () => {
    setExitSaving(true)
    try {
      const ok = await save(true)
      if (!ok) return
      setExitConfirmOpen(false)
      await window.slidesApi.requestHostClose?.()
    } finally {
      setExitSaving(false)
    }
  }, [save])

'''
text = must_replace(text, marker, marker + lifecycle, 'web lifecycle callbacks')
# Pass Ribbon optional callbacks just after onSaveAs.
text = must_replace(
    text,
    "        onSaveAs={() => void saveAs()}\n        onExportPdf={() => void exportPdf()}\n",
    "        onSaveAs={() => void saveAs()}\n        onSaveHistoryVersion={window.slidesApi.saveHistoryVersion ? () => void saveHistoryVersion() : undefined}\n        onExportPptx={window.slidesApi.exportPptx ? () => void exportPptx() : undefined}\n        onExit={window.slidesApi.requestHostClose ? () => void requestExit() : undefined}\n        onExportPdf={() => void exportPdf()}\n",
    'ribbon lifecycle props',
)
# Exit modal before context menu near bottom.
marker = "      {ctxMenu && (\n"
modal = '''      {exitConfirmOpen && (
        <div className="modal-backdrop" onClick={() => !exitSaving && setExitConfirmOpen(false)}>
          <div className="modal" onClick={(event) => event.stopPropagation()}>
            <h2>{webLabels.exitTitle}</h2>
            <p>{webLabels.exitMessage}</p>
            <div className="modal-actions">
              <button disabled={exitSaving} onClick={() => setExitConfirmOpen(false)}>
                {webLabels.cancel}
              </button>
              <button
                disabled={exitSaving}
                onClick={() => {
                  setExitConfirmOpen(false)
                  void window.slidesApi.requestHostClose?.()
                }}
              >
                {webLabels.discardAndExit}
              </button>
              <button className="primary" disabled={exitSaving} onClick={() => void saveAndExit()}>
                {webLabels.saveAndExit}
              </button>
            </div>
          </div>
        </div>
      )}

'''
text = must_replace(text, marker, modal + marker, 'exit modal')
write(path, text)

print('PPT Web patch applied')
