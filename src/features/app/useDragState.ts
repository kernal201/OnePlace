import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react'
import type { DragPosition, DragState, DropTarget } from '../../app/appModel'

export const useDragState = () => {
  const [dragState, setDragState] = useState<DragState | null>(null)
  const [dropTarget, setDropTarget] = useState<DropTarget | null>(null)
  const [dragPosition, setDragPosition] = useState<DragPosition | null>(null)
  const suppressClickAfterDragRef = useRef(false)

  useEffect(() => {
    if (!dragState) return

    const syncPointer = (event: PointerEvent) => {
      setDragPosition({ x: event.clientX, y: event.clientY })
    }

    const stopDragging = () => {
      setDragState(null)
      setDragPosition(null)
      setDropTarget(null)
    }

    window.addEventListener('pointermove', syncPointer)
    window.addEventListener('pointerup', stopDragging)
    return () => {
      window.removeEventListener('pointermove', syncPointer)
      window.removeEventListener('pointerup', stopDragging)
    }
  }, [dragState])

  const clearDragState = (suppressClick = false) => {
    if (suppressClick) {
      suppressClickAfterDragRef.current = true
    }
    setDragState(null)
    setDropTarget(null)
    setDragPosition(null)
  }

  const consumeSuppressedClick = () => {
    if (!suppressClickAfterDragRef.current) return false
    suppressClickAfterDragRef.current = false
    return true
  }

  const beginDrag = (event: ReactPointerEvent<HTMLElement>, nextState: DragState) => {
    if (event.button !== 0) return
    suppressClickAfterDragRef.current = false
    setDragState(nextState)
    setDropTarget(null)
    setDragPosition({ x: event.clientX, y: event.clientY })
  }

  const allowDrop = (event: ReactPointerEvent<HTMLElement>) => {
    if (!dragState) return
    event.preventDefault()
  }

  return {
    allowDrop,
    beginDrag,
    clearDragState,
    consumeSuppressedClick,
    dragPosition,
    dragState,
    dropTarget,
    setDropTarget,
  }
}
