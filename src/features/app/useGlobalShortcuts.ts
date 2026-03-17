import { useEffect, type RefObject } from 'react'

type ShortcutActions = {
  addPage: () => void
  createNotebook: () => void
  demoteCurrentPage: () => void
  openNotebook: () => void
  promoteCurrentPage: () => void
  saveNow: () => Promise<void> | void
  saveNotebookAs: () => Promise<void> | void
}

type UseGlobalShortcutsArgs = {
  canDemotePage: boolean
  canPromotePage: boolean
  searchInputRef: RefObject<HTMLInputElement | null>
  shortcutActions: ShortcutActions
}

export const useGlobalShortcuts = ({
  canDemotePage,
  canPromotePage,
  searchInputRef,
  shortcutActions,
}: UseGlobalShortcutsArgs) => {
  useEffect(() => {
    const handleShortcut = (event: KeyboardEvent) => {
      const target = event.target
      const isEditable =
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        (target instanceof HTMLElement && target.isContentEditable)

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault()
        void shortcutActions.saveNow()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
        event.preventDefault()
        searchInputRef.current?.focus()
        searchInputRef.current?.select()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'n' && !event.shiftKey) {
        event.preventDefault()
        shortcutActions.addPage()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 'n') {
        event.preventDefault()
        shortcutActions.createNotebook()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'o') {
        event.preventDefault()
        shortcutActions.openNotebook()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 's') {
        event.preventDefault()
        void shortcutActions.saveNotebookAs()
        return
      }

      if (isEditable) return

      if ((event.ctrlKey || event.metaKey) && event.key === 'ArrowUp' && canPromotePage) {
        event.preventDefault()
        shortcutActions.promoteCurrentPage()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key === 'ArrowDown' && canDemotePage) {
        event.preventDefault()
        shortcutActions.demoteCurrentPage()
      }
    }

    window.addEventListener('keydown', handleShortcut)
    return () => {
      window.removeEventListener('keydown', handleShortcut)
    }
  }, [canDemotePage, canPromotePage, searchInputRef, shortcutActions])
}
