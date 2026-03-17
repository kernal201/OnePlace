import type { ContextMenuState } from '../../app/appModel'

type ContextMenuOverlayProps = {
  contextMenu: ContextMenuState | null
  onRequestClose: () => void
}

export function ContextMenuOverlay({ contextMenu, onRequestClose }: ContextMenuOverlayProps) {
  if (!contextMenu) return null

  return (
    <div
      className="context-menu"
      style={{ left: `${contextMenu.x}px`, top: `${contextMenu.y}px` }}
    >
      {contextMenu.items.map((item) => (
        <button
          key={item.label}
          className={`context-menu-item ${item.danger ? 'danger' : ''}`}
          disabled={item.disabled}
          onClick={() => {
            onRequestClose()
            item.onSelect()
          }}
          type="button"
        >
          {item.label}
        </button>
      ))}
    </div>
  )
}
