import { useEffect, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { Lang } from '@genoffice/i18n'
import { useI18n } from '../i18n/locale'

type ExitLabels = {
  title: string
  message: string
  cancel: string
  discard: string
  save: string
}

const EXIT_LABELS: Record<Lang, ExitLabels> = {
  zh: {
    title: '退出前保存更改?',
    message: '此文档有尚未保存的更改。',
    cancel: '取消',
    discard: '放弃并退出',
    save: '保存并退出',
  },
  'zh-TW': {
    title: '結束前儲存變更？',
    message: '此文件有尚未儲存的變更。',
    cancel: '取消',
    discard: '放棄並結束',
    save: '儲存並結束',
  },
  en: {
    title: 'Save changes before exiting?',
    message: 'This document has unsaved changes.',
    cancel: 'Cancel',
    discard: 'Discard and Exit',
    save: 'Save and Exit',
  },
  ja: {
    title: '終了する前に変更を保存しますか？',
    message: 'この文書には保存されていない変更があります。',
    cancel: 'キャンセル',
    discard: '破棄して終了',
    save: '保存して終了',
  },
  ko: {
    title: '종료하기 전에 변경 내용을 저장할까요?',
    message: '이 문서에 저장되지 않은 변경 내용이 있습니다.',
    cancel: '취소',
    discard: '저장 안 함 및 종료',
    save: '저장 후 종료',
  },
  fr: {
    title: 'Enregistrer les modifications avant de quitter ?',
    message: 'Ce document contient des modifications non enregistrées.',
    cancel: 'Annuler',
    discard: 'Ignorer et quitter',
    save: 'Enregistrer et quitter',
  },
  de: {
    title: 'Änderungen vor dem Beenden speichern?',
    message: 'Dieses Dokument enthält nicht gespeicherte Änderungen.',
    cancel: 'Abbrechen',
    discard: 'Verwerfen und beenden',
    save: 'Speichern und beenden',
  },
  es: {
    title: '¿Guardar los cambios antes de salir?',
    message: 'Este documento tiene cambios sin guardar.',
    cancel: 'Cancelar',
    discard: 'Descartar y salir',
    save: 'Guardar y salir',
  },
  th: {
    title: 'บันทึกการเปลี่ยนแปลงก่อนออกหรือไม่',
    message: 'เอกสารนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    cancel: 'ยกเลิก',
    discard: 'ละทิ้งและออก',
    save: 'บันทึกและออก',
  },
  id: {
    title: 'Simpan perubahan sebelum keluar?',
    message: 'Dokumen ini memiliki perubahan yang belum disimpan.',
    cancel: 'Batal',
    discard: 'Buang dan keluar',
    save: 'Simpan dan keluar',
  },
  ru: {
    title: 'Сохранить изменения перед выходом?',
    message: 'В этом документе есть несохраненные изменения.',
    cancel: 'Отмена',
    discard: 'Не сохранять и выйти',
    save: 'Сохранить и выйти',
  },
  ar: {
    title: 'هل تريد حفظ التغييرات قبل الخروج؟',
    message: 'يحتوي هذا المستند على تغييرات غير محفوظة.',
    cancel: 'إلغاء',
    discard: 'تجاهل والخروج',
    save: 'حفظ والخروج',
  },
  pt: {
    title: 'Salvar alterações antes de sair?',
    message: 'Este documento tem alterações não salvas.',
    cancel: 'Cancelar',
    discard: 'Descartar e sair',
    save: 'Salvar e sair',
  },
  it: {
    title: 'Salvare le modifiche prima di uscire?',
    message: 'Questo documento contiene modifiche non salvate.',
    cancel: 'Annulla',
    discard: 'Ignora ed esci',
    save: 'Salva ed esci',
  },
  pl: {
    title: 'Zapisać zmiany przed wyjściem?',
    message: 'Ten dokument zawiera niezapisane zmiany.',
    cancel: 'Anuluj',
    discard: 'Odrzuć i wyjdź',
    save: 'Zapisz i wyjdź',
  },
  nl: {
    title: 'Wijzigingen opslaan voordat u afsluit?',
    message: 'Dit document bevat niet-opgeslagen wijzigingen.',
    cancel: 'Annuleren',
    discard: 'Negeren en afsluiten',
    save: 'Opslaan en afsluiten',
  },
  ms: {
    title: 'Simpan perubahan sebelum keluar?',
    message: 'Dokumen ini mempunyai perubahan yang belum disimpan.',
    cancel: 'Batal',
    discard: 'Buang dan keluar',
    save: 'Simpan dan keluar',
  },
  he: {
    title: 'לשמור שינויים לפני היציאה?',
    message: 'במסמך זה יש שינויים שלא נשמרו.',
    cancel: 'ביטול',
    discard: 'בטל שינויים וצא',
    save: 'שמור וצא',
  },
  hi: {
    title: 'बाहर निकलने से पहले बदलाव सहेजें?',
    message: 'इस दस्तावेज़ में बिना सहेजे बदलाव हैं।',
    cancel: 'रद्द करें',
    discard: 'बदलाव छोड़ें और बाहर निकलें',
    save: 'सहेजें और बाहर निकलें',
  },
}

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
  const { lang } = useI18n()
  const labels = EXIT_LABELS[lang]
  const backdropRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const root = backdropRef.current
    if (!root || root.contains(document.activeElement)) return
    root.querySelector<HTMLElement>('button')?.focus()
  }, [])

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key !== 'Escape') return
    event.preventDefault()
    event.stopPropagation()
    onCancel()
  }

  return (
    <div
      className="modal-backdrop"
      ref={backdropRef}
      onKeyDown={onKeyDown}
      onMouseDown={(event) => event.target === event.currentTarget && !saving && onCancel()}
    >
      <div className="modal" role="dialog" aria-modal="true" aria-labelledby="markdown-exit-title">
        <h2 id="markdown-exit-title">{labels.title}</h2>
        <p>{labels.message}</p>
        <div className="modal-actions">
          <button className="btn-ghost" disabled={saving} onClick={onCancel}>
            {labels.cancel}
          </button>
          <button className="btn-ghost" disabled={saving} onClick={onDiscard}>
            {labels.discard}
          </button>
          <button className="btn-primary" disabled={saving} onClick={onSave}>
            {labels.save}
          </button>
        </div>
      </div>
    </div>
  )
}
