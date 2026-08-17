import { useEffect, useRef, useState } from 'react'
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
  fr: {
    saveHistoryVersion: 'Enregistrer une nouvelle version',
    download: 'Télécharger',
    exit: 'Quitter',
  },
  de: {
    saveHistoryVersion: 'Als neue Verlaufsversion speichern',
    download: 'Herunterladen',
    exit: 'Beenden',
  },
  es: {
    saveHistoryVersion: 'Guardar nueva versión del historial',
    download: 'Descargar',
    exit: 'Salir',
  },
  th: { saveHistoryVersion: 'บันทึกเป็นเวอร์ชันประวัติใหม่', download: 'ดาวน์โหลด', exit: 'ออก' },
  id: {
    saveHistoryVersion: 'Simpan sebagai versi riwayat baru',
    download: 'Unduh',
    exit: 'Keluar',
  },
  ru: { saveHistoryVersion: 'Сохранить новую версию истории', download: 'Скачать', exit: 'Выйти' },
  ar: { saveHistoryVersion: 'حفظ كإصدار سجل جديد', download: 'تنزيل', exit: 'خروج' },
  pt: { saveHistoryVersion: 'Salvar nova versão do histórico', download: 'Baixar', exit: 'Sair' },
  it: { saveHistoryVersion: 'Salva nuova versione cronologia', download: 'Scarica', exit: 'Esci' },
  pl: { saveHistoryVersion: 'Zapisz nową wersję historii', download: 'Pobierz', exit: 'Wyjdź' },
  nl: {
    saveHistoryVersion: 'Opslaan als nieuwe geschiedenisversie',
    download: 'Downloaden',
    exit: 'Afsluiten',
  },
  ms: {
    saveHistoryVersion: 'Simpan sebagai versi sejarah baharu',
    download: 'Muat turun',
    exit: 'Keluar',
  },
  he: { saveHistoryVersion: 'שמירה כגרסת היסטוריה חדשה', download: 'הורדה', exit: 'יציאה' },
  hi: {
    saveHistoryVersion: 'नए इतिहास संस्करण के रूप में सहेजें',
    download: 'डाउनलोड',
    exit: 'बाहर निकलें',
  },
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
