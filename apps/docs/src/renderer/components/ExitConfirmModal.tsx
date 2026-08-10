import { useI18n } from '../i18n/locale'
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
