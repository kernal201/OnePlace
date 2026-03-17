import type { DragPosition, DragState } from '../../app/appModel'

type DragBadgeProps = {
  dragState: DragState | null
  dragPosition: DragPosition | null
  dragLabel: string
}

export function DragBadge({ dragState, dragPosition, dragLabel }: DragBadgeProps) {
  if (!dragState || !dragPosition) return null

  return (
    <div
      className="drag-badge"
      style={{ left: `${dragPosition.x + 18}px`, top: `${dragPosition.y + 18}px` }}
    >
      {dragLabel}
    </div>
  )
}
