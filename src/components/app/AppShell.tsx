import type { ComponentProps } from 'react'
import type { ContextMenuState, DragPosition, DragState } from '../../app/appModel'
import { ContextMenuOverlay } from './ContextMenuOverlay'
import { DragBadge } from './DragBadge'
import { FileInputs } from './FileInputs'
import { NotePane } from './NotePane'
import { NotebooksPane } from './NotebooksPane'
import { PagesPane } from './PagesPane'
import { RibbonBar } from './RibbonBar'
import { TitleBar } from './TitleBar'

type AppShellProps = {
  contextMenu: ContextMenuState | null
  dragLabel: string
  dragPosition: DragPosition | null
  dragState: DragState | null
  fileInputProps: ComponentProps<typeof FileInputs>
  isNotebookPaneVisible: boolean
  isPagesPaneVisible: boolean
  notePaneProps: ComponentProps<typeof NotePane>
  notebooksPaneProps: ComponentProps<typeof NotebooksPane>
  onRequestCloseContextMenu: () => void
  pagesPaneProps: ComponentProps<typeof PagesPane>
  ribbonBarProps: ComponentProps<typeof RibbonBar>
  titleBarProps: ComponentProps<typeof TitleBar>
}

export function AppShell({
  contextMenu,
  dragLabel,
  dragPosition,
  dragState,
  fileInputProps,
  isNotebookPaneVisible,
  isPagesPaneVisible,
  notePaneProps,
  notebooksPaneProps,
  onRequestCloseContextMenu,
  pagesPaneProps,
  ribbonBarProps,
  titleBarProps,
}: AppShellProps) {
  return (
    <div className={`desktop-scene ${dragState ? 'drag-active' : ''}`}>
      <div className={`onenote-window ${dragState ? `dragging-${dragState.type}` : ''}`}>
        <TitleBar {...titleBarProps} />
        <RibbonBar {...ribbonBarProps} />
        <main
          className={`workspace ${isNotebookPaneVisible ? '' : 'notebooks-hidden'} ${isPagesPaneVisible ? '' : 'pages-hidden'}`}
        >
          <NotebooksPane {...notebooksPaneProps} />
          <PagesPane {...pagesPaneProps} />
          <NotePane {...notePaneProps} />
        </main>
      </div>
      <FileInputs {...fileInputProps} />
      <ContextMenuOverlay contextMenu={contextMenu} onRequestClose={onRequestCloseContextMenu} />
      <DragBadge dragLabel={dragLabel} dragPosition={dragPosition} dragState={dragState} />
    </div>
  )
}
