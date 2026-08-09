import type { Lang } from '@genoffice/i18n'

export interface SlidesWebLifecycleLabels {
  saveHistory: string
  exportPptx: string
  exit: string
  historySaved: string
  historyFailed: string
  exportDone: string
  exportFailed: string
  exitTitle: string
  exitMessage: string
  cancel: string
  discardAndExit: string
  saveAndExit: string
}

const en: SlidesWebLifecycleLabels = {
  saveHistory: 'Save as History Version',
  exportPptx: 'Export as PPTX',
  exit: 'Exit',
  historySaved: 'History version saved',
  historyFailed: 'Failed to save history version',
  exportDone: 'PPTX exported',
  exportFailed: 'PPTX export failed',
  exitTitle: 'Save changes before exiting?',
  exitMessage: 'This presentation has unsaved changes.',
  cancel: 'Cancel',
  discardAndExit: 'Discard and Exit',
  saveAndExit: 'Save and Exit',
}

const labels: Partial<Record<Lang, SlidesWebLifecycleLabels>> = {
  en,
  zh: {
    saveHistory: '存为历史版本',
    exportPptx: '导出为 PPTX',
    exit: '退出',
    historySaved: '已存为历史版本',
    historyFailed: '存为历史版本失败',
    exportDone: 'PPTX 已导出',
    exportFailed: 'PPTX 导出失败',
    exitTitle: '退出前保存更改？',
    exitMessage: '此演示文稿有尚未保存的更改。',
    cancel: '取消',
    discardAndExit: '放弃并退出',
    saveAndExit: '保存并退出',
  },
  'zh-TW': {
    saveHistory: '另存為歷史版本',
    exportPptx: '匯出為 PPTX',
    exit: '退出',
    historySaved: '已儲存歷史版本',
    historyFailed: '儲存歷史版本失敗',
    exportDone: 'PPTX 已匯出',
    exportFailed: 'PPTX 匯出失敗',
    exitTitle: '退出前儲存變更？',
    exitMessage: '此簡報有尚未儲存的變更。',
    cancel: '取消',
    discardAndExit: '放棄並退出',
    saveAndExit: '儲存並退出',
  },
  pt: {
    saveHistory: 'Salvar como versão histórica',
    exportPptx: 'Exportar como PPTX',
    exit: 'Sair',
    historySaved: 'Versão histórica salva',
    historyFailed: 'Falha ao salvar a versão histórica',
    exportDone: 'PPTX exportado',
    exportFailed: 'Falha ao exportar PPTX',
    exitTitle: 'Salvar alterações antes de sair?',
    exitMessage: 'Esta apresentação tem alterações não salvas.',
    cancel: 'Cancelar',
    discardAndExit: 'Descartar e sair',
    saveAndExit: 'Salvar e sair',
  },
  ja: {
    ...en,
    saveHistory: '履歴バージョンとして保存',
    exportPptx: 'PPTXとしてエクスポート',
    exit: '終了',
    cancel: 'キャンセル',
    discardAndExit: '保存せず終了',
    saveAndExit: '保存して終了',
  },
  ko: {
    ...en,
    saveHistory: '기록 버전으로 저장',
    exportPptx: 'PPTX로 내보내기',
    exit: '종료',
    cancel: '취소',
    discardAndExit: '저장하지 않고 종료',
    saveAndExit: '저장 후 종료',
  },
  fr: {
    ...en,
    saveHistory: 'Enregistrer comme version historique',
    exportPptx: 'Exporter en PPTX',
    exit: 'Quitter',
    cancel: 'Annuler',
    discardAndExit: 'Ignorer et quitter',
    saveAndExit: 'Enregistrer et quitter',
  },
  de: {
    ...en,
    saveHistory: 'Als Verlaufsversion speichern',
    exportPptx: 'Als PPTX exportieren',
    exit: 'Beenden',
    cancel: 'Abbrechen',
    discardAndExit: 'Verwerfen und beenden',
    saveAndExit: 'Speichern und beenden',
  },
  es: {
    ...en,
    saveHistory: 'Guardar como versión histórica',
    exportPptx: 'Exportar como PPTX',
    exit: 'Salir',
    cancel: 'Cancelar',
    discardAndExit: 'Descartar y salir',
    saveAndExit: 'Guardar y salir',
  },
  it: {
    ...en,
    saveHistory: 'Salva come versione storica',
    exportPptx: 'Esporta come PPTX',
    exit: 'Esci',
    cancel: 'Annulla',
    discardAndExit: 'Ignora ed esci',
    saveAndExit: 'Salva ed esci',
  },
  pl: {
    ...en,
    saveHistory: 'Zapisz jako wersję historyczną',
    exportPptx: 'Eksportuj jako PPTX',
    exit: 'Zakończ',
    cancel: 'Anuluj',
    discardAndExit: 'Odrzuć i zakończ',
    saveAndExit: 'Zapisz i zakończ',
  },
  nl: {
    ...en,
    saveHistory: 'Opslaan als historische versie',
    exportPptx: 'Exporteren als PPTX',
    exit: 'Afsluiten',
    cancel: 'Annuleren',
    discardAndExit: 'Negeren en afsluiten',
    saveAndExit: 'Opslaan en afsluiten',
  },
  th: {
    ...en,
    saveHistory: 'บันทึกเป็นเวอร์ชันประวัติ',
    exportPptx: 'ส่งออกเป็น PPTX',
    exit: 'ออก',
    cancel: 'ยกเลิก',
  },
  id: {
    ...en,
    saveHistory: 'Simpan sebagai versi riwayat',
    exportPptx: 'Ekspor sebagai PPTX',
    exit: 'Keluar',
    cancel: 'Batal',
  },
  ru: {
    ...en,
    saveHistory: 'Сохранить как историческую версию',
    exportPptx: 'Экспортировать в PPTX',
    exit: 'Выйти',
    cancel: 'Отмена',
  },
  ar: {
    ...en,
    saveHistory: 'حفظ كإصدار محفوظ',
    exportPptx: 'تصدير كـ PPTX',
    exit: 'خروج',
    cancel: 'إلغاء',
  },
  ms: {
    ...en,
    saveHistory: 'Simpan sebagai versi sejarah',
    exportPptx: 'Eksport sebagai PPTX',
    exit: 'Keluar',
    cancel: 'Batal',
  },
  he: {
    ...en,
    saveHistory: 'שמירה כגרסת היסטוריה',
    exportPptx: 'ייצוא כ-PPTX',
    exit: 'יציאה',
    cancel: 'ביטול',
  },
  hi: {
    ...en,
    saveHistory: 'इतिहास संस्करण के रूप में सहेजें',
    exportPptx: 'PPTX के रूप में निर्यात करें',
    exit: 'बाहर निकलें',
    cancel: 'रद्द करें',
  },
}

export function slidesWebLifecycleLabels(lang: Lang): SlidesWebLifecycleLabels {
  return labels[lang] ?? en
}
