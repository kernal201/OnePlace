import { useEffect, useState, type MouseEvent as ReactMouseEvent } from 'react'
import type { ContextMenuItem, ContextMenuState } from '../../app/appModel'

export const useContextMenu = () => {
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null)

  useEffect(() => {
    if (!contextMenu) return

    const dismiss = () => setContextMenu(null)
    const dismissOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') dismiss()
    }

    window.addEventListener('click', dismiss)
    window.addEventListener('contextmenu', dismiss)
    window.addEventListener('keydown', dismissOnEscape)
    window.addEventListener('resize', dismiss)
    return () => {
      window.removeEventListener('click', dismiss)
      window.removeEventListener('contextmenu', dismiss)
      window.removeEventListener('keydown', dismissOnEscape)
      window.removeEventListener('resize', dismiss)
    }
  }, [contextMenu])

  const openContextMenu = (event: ReactMouseEvent<HTMLElement>, items: ContextMenuItem[]) => {
    event.preventDefault()
    setContextMenu({
      items,
      x: event.clientX,
      y: event.clientY,
    })
  }

  return {
    contextMenu,
    openContextMenu,
    setContextMenu,
  }
}
