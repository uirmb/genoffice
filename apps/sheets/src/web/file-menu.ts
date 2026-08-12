import { t } from '../renderer/i18n/locale'
import { dispatchSheetsWebFileAction, type SheetsWebFileAction } from './file-actions'

interface FileMenuElements {
  root: HTMLDivElement
  trigger: HTMLButtonElement
  menu: HTMLDivElement
  open: HTMLButtonElement
  save: HTMLButtonElement
  saveAs: HTMLButtonElement
  saveHistory: HTMLButtonElement
  exportXlsx: HTMLButtonElement
  exit: HTMLButtonElement
}

function button(className: string): HTMLButtonElement {
  const element = document.createElement('button')
  element.type = 'button'
  element.className = className
  return element
}

function setMenuLabel(
  element: HTMLButtonElement,
  label: string,
  shortcut?: string,
): void {
  element.replaceChildren()
  const text = document.createElement('span')
  text.textContent = label
  element.append(text)
  if (shortcut) {
    const key = document.createElement('span')
    key.className = 'file-menu-key'
    key.textContent = shortcut
    element.append(key)
  }
}

function currentWorkbookDirty(): boolean {
  const saveButton = document.querySelector<HTMLButtonElement>('.ribbon-tabs .qa-btn')
  return saveButton ? !saveButton.disabled : false
}

function createExitDialog(onClose: () => void, onSaveAndExit: () => void): HTMLDivElement {
  const backdrop = document.createElement('div')
  backdrop.className = 'file-exit-backdrop'

  const dialog = document.createElement('div')
  dialog.className = 'file-exit-dialog'
  dialog.setAttribute('role', 'dialog')
  dialog.setAttribute('aria-modal', 'true')

  const heading = document.createElement('h2')
  heading.textContent = t('appFileUnsavedTitle')
  const message = document.createElement('p')
  message.textContent = t('appFileUnsavedMessage')
  const actions = document.createElement('div')
  actions.className = 'file-exit-actions'

  const cancel = button('')
  cancel.textContent = t('appFileCancel')
  cancel.addEventListener('click', onClose)

  const discard = button('')
  discard.textContent = t('appFileDiscardAndExit')
  discard.addEventListener('click', () => {
    onClose()
    dispatchSheetsWebFileAction('discard-and-exit')
  })

  const save = button('primary')
  save.textContent = t('appFileSaveAndExit')
  save.addEventListener('click', () => {
    onClose()
    onSaveAndExit()
  })

  actions.append(cancel, discard, save)
  dialog.append(heading, message, actions)
  backdrop.append(dialog)
  backdrop.addEventListener('mousedown', (event) => {
    if (event.target === backdrop) onClose()
  })
  return backdrop
}

function createFileMenu(): FileMenuElements {
  const root = document.createElement('div')
  root.className = 'sheets-web-file-menu-root file-tab-wrap'

  const trigger = button('ribbon-tab ribbon-tab-file')
  trigger.setAttribute('aria-haspopup', 'menu')
  trigger.setAttribute('aria-expanded', 'false')

  const menu = document.createElement('div')
  menu.className = 'file-menu'
  menu.setAttribute('role', 'menu')
  menu.hidden = true

  const open = button('file-menu-open')
  const save = button('file-menu-save')
  const saveAs = button('file-menu-save-as')
  const saveHistory = button('file-menu-save-history')
  const exportXlsx = button('file-menu-export-xlsx')
  const exit = button('file-menu-exit')

  menu.append(open, save, saveAs, saveHistory, exportXlsx, exit)
  root.append(trigger, menu)
  return { root, trigger, menu, open, save, saveAs, saveHistory, exportXlsx, exit }
}

export function installSheetsWebFileMenu(): () => void {
  const elements = createFileMenu()
  document.body.append(elements.root)
  let exitDialog: HTMLDivElement | null = null
  let saveExitTimer: ReturnType<typeof setTimeout> | null = null

  const updateLabels = (): void => {
    elements.trigger.textContent = t('appFileTab')
    setMenuLabel(elements.open, t('appFileOpen'), 'Ctrl+O')
    setMenuLabel(elements.save, t('appFileSave'), 'Ctrl+S')
    setMenuLabel(elements.saveAs, t('appFileSaveAs'), 'Ctrl+Shift+S')
    setMenuLabel(elements.saveHistory, t('appFileSaveHistory'))
    setMenuLabel(elements.exportXlsx, t('appFileExportXlsx'))
    setMenuLabel(elements.exit, t('appFileExit'))
  }

  const closeMenu = (): void => {
    elements.menu.hidden = true
    elements.trigger.classList.remove('open')
    elements.trigger.setAttribute('aria-expanded', 'false')
  }

  const openMenu = (): void => {
    const readOnly = document.documentElement.dataset.officeMode === 'view'
    elements.save.disabled = readOnly || !currentWorkbookDirty()
    elements.saveAs.disabled = readOnly
    elements.saveHistory.disabled = readOnly
    elements.menu.hidden = false
    elements.trigger.classList.add('open')
    elements.trigger.setAttribute('aria-expanded', 'true')
  }

  const toggleMenu = (): void => {
    if (elements.menu.hidden) openMenu()
    else closeMenu()
  }

  const run = (action: SheetsWebFileAction): void => {
    closeMenu()
    dispatchSheetsWebFileAction(action)
  }

  const closeExitDialog = (): void => {
    exitDialog?.remove()
    exitDialog = null
  }

  const saveAndExit = (): void => {
    if (saveExitTimer !== null) clearTimeout(saveExitTimer)
    dispatchSheetsWebFileAction('save')
    const deadline = Date.now() + 120_000
    const waitForClean = (): void => {
      if (!currentWorkbookDirty()) {
        saveExitTimer = null
        dispatchSheetsWebFileAction('discard-and-exit')
        return
      }
      // Save failures and cancellations leave the journal dirty. In that case
      // this one-shot close attempt simply expires; a later ordinary Save can
      // never inherit a stale "close after save" flag.
      if (Date.now() >= deadline) {
        saveExitTimer = null
        return
      }
      saveExitTimer = setTimeout(waitForClean, 100)
    }
    saveExitTimer = setTimeout(waitForClean, 100)
  }

  const requestExit = (): void => {
    closeMenu()
    if (!currentWorkbookDirty()) {
      dispatchSheetsWebFileAction('discard-and-exit')
      return
    }
    closeExitDialog()
    exitDialog = createExitDialog(closeExitDialog, saveAndExit)
    document.body.append(exitDialog)
  }

  elements.trigger.addEventListener('click', toggleMenu)
  elements.open.addEventListener('click', () => run('open'))
  elements.save.addEventListener('click', () => run('save'))
  elements.saveAs.addEventListener('click', () => run('save-as'))
  elements.saveHistory.addEventListener('click', () => run('save-history'))
  elements.exportXlsx.addEventListener('click', () => run('export-xlsx'))
  elements.exit.addEventListener('click', requestExit)

  const onDocumentPointerDown = (event: MouseEvent): void => {
    if (!elements.root.contains(event.target as Node)) closeMenu()
  }
  const onDocumentKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      closeMenu()
      closeExitDialog()
    }
  }
  document.addEventListener('mousedown', onDocumentPointerDown)
  document.addEventListener('keydown', onDocumentKeyDown)

  updateLabels()
  const unsubscribeLanguage = window.desktopApi.onLanguageChanged(() => updateLabels())

  return () => {
    unsubscribeLanguage()
    document.removeEventListener('mousedown', onDocumentPointerDown)
    document.removeEventListener('keydown', onDocumentKeyDown)
    if (saveExitTimer !== null) clearTimeout(saveExitTimer)
    closeExitDialog()
    elements.root.remove()
  }
}
