import { useCallback, useEffect, type MutableRefObject, type RefObject } from 'react'
import { dehydratePageContent, hydratePageContent, normalizeTerminalText, sanitizePastedHtml } from '../../app/appModel'
import type { AppState, Page, PageUpdate } from '../../app/appModel'

type UseEditorChromeArgs = {
  appState: AppState
  page: Page | undefined
  editorRef: RefObject<HTMLDivElement | null>
  noteCanvasScrollRef: RefObject<HTMLDivElement | null>
  titlebarSearchRef: RefObject<HTMLDivElement | null>
  searchResultsPanelRef: RefObject<HTMLDivElement | null>
  styleMenuRef: RefObject<HTMLDivElement | null>
  styleMenuPanelRef: RefObject<HTMLDivElement | null>
  fontMenuRef: RefObject<HTMLDivElement | null>
  fontMenuPanelRef: RefObject<HTMLDivElement | null>
  fontSizeMenuRef: RefObject<HTMLDivElement | null>
  fontSizeMenuPanelRef: RefObject<HTMLDivElement | null>
  caretScrollFrameRef: MutableRefObject<number | null>
  lastHydratedPageIdRef: MutableRefObject<string | null>
  selectionRangeRef: MutableRefObject<Range | null>
  suppressAutoFollowUntilRef: MutableRefObject<number>
  setIsFontMenuOpen: (value: boolean) => void
  setIsFontSizeMenuOpen: (value: boolean) => void
  setIsStyleMenuOpen: (value: boolean) => void
  setQuery: (value: string) => void
  updatePage: (updates: PageUpdate) => void
}

export const useEditorChrome = ({
  appState,
  page,
  editorRef,
  noteCanvasScrollRef,
  titlebarSearchRef,
  searchResultsPanelRef,
  styleMenuRef,
  styleMenuPanelRef,
  fontMenuRef,
  fontMenuPanelRef,
  fontSizeMenuRef,
  fontSizeMenuPanelRef,
  caretScrollFrameRef,
  lastHydratedPageIdRef,
  selectionRangeRef,
  suppressAutoFollowUntilRef,
  setIsFontMenuOpen,
  setIsFontSizeMenuOpen,
  setIsStyleMenuOpen,
  setQuery,
  updatePage,
}: UseEditorChromeArgs) => {
  useEffect(() => {
    return () => {
      if (caretScrollFrameRef.current) {
        window.cancelAnimationFrame(caretScrollFrameRef.current)
      }
    }
  }, [caretScrollFrameRef])

  const pauseAutoFollow = useCallback(
    (durationMs = 900) => {
      suppressAutoFollowUntilRef.current = Date.now() + durationMs
    },
    [suppressAutoFollowUntilRef],
  )

  const keepCaretInView = useCallback(
    (behavior: ScrollBehavior = 'auto') => {
      if (caretScrollFrameRef.current) {
        window.cancelAnimationFrame(caretScrollFrameRef.current)
      }

      caretScrollFrameRef.current = window.requestAnimationFrame(() => {
        caretScrollFrameRef.current = null

        const shell = noteCanvasScrollRef.current
        const editor = editorRef.current
        const selection = window.getSelection()
        if (!shell || !editor || !selection || selection.rangeCount === 0) return
        if (!selection.isCollapsed) return
        if (Date.now() < suppressAutoFollowUntilRef.current) return

        const range = selection.getRangeAt(0)
        if (!editor.contains(range.commonAncestorContainer)) return

        const caretRange = range.cloneRange()
        caretRange.collapse(false)
        let caretRect = caretRange.getBoundingClientRect()

        if (caretRect.width === 0 && caretRect.height === 0) {
          const marker = document.createElement('span')
          marker.textContent = '\u200b'
          caretRange.insertNode(marker)
          caretRect = marker.getBoundingClientRect()
          marker.parentNode?.removeChild(marker)
          selection.removeAllRanges()
          selection.addRange(range)
        }

        const shellRect = shell.getBoundingClientRect()
        const topComfortLine = shellRect.top + Math.min(88, shellRect.height * 0.18)
        const lowerComfortLine = shellRect.top + shellRect.height * 0.68
        let nextScrollTop = shell.scrollTop

        if (caretRect.bottom > lowerComfortLine) {
          nextScrollTop += caretRect.bottom - lowerComfortLine
        } else if (caretRect.top < topComfortLine) {
          nextScrollTop -= topComfortLine - caretRect.top
        }

        nextScrollTop = Math.max(0, Math.min(nextScrollTop, shell.scrollHeight - shell.clientHeight))
        if (Math.abs(nextScrollTop - shell.scrollTop) < 2) return

        shell.scrollTo({ top: nextScrollTop, behavior })
      })
    },
    [caretScrollFrameRef, editorRef, noteCanvasScrollRef, suppressAutoFollowUntilRef],
  )

  useEffect(() => {
    const shell = noteCanvasScrollRef.current
    if (!shell) return

    const handleWheel = () => pauseAutoFollow(1200)
    const handlePointerDown = () => pauseAutoFollow(500)
    const handleTouchMove = () => pauseAutoFollow(1200)

    shell.addEventListener('wheel', handleWheel, { passive: true })
    shell.addEventListener('pointerdown', handlePointerDown)
    shell.addEventListener('touchmove', handleTouchMove, { passive: true })
    return () => {
      shell.removeEventListener('wheel', handleWheel)
      shell.removeEventListener('pointerdown', handlePointerDown)
      shell.removeEventListener('touchmove', handleTouchMove)
    }
  }, [noteCanvasScrollRef, page?.id, pauseAutoFollow])

  useEffect(() => {
    const editor = editorRef.current
    if (!editor || !page) return

    const pageChanged = lastHydratedPageIdRef.current !== page.id
    const isFocused = document.activeElement === editor

    if (!pageChanged && isFocused) return
    if (pageChanged) {
      selectionRangeRef.current = null
    }

    const hydrated = hydratePageContent(page.content, appState.meta.assets)
    if (editor.innerHTML !== hydrated) {
      editor.innerHTML = hydrated
    }

    lastHydratedPageIdRef.current = page.id

    if (pageChanged && isFocused) {
      window.setTimeout(() => {
        if (!editorRef.current || document.activeElement !== editorRef.current) return
        const range = document.createRange()
        range.selectNodeContents(editorRef.current)
        range.collapse(false)
        const selection = window.getSelection()
        selection?.removeAllRanges()
        selection?.addRange(range)
        selectionRangeRef.current = range.cloneRange()
      }, 0)
    }
  }, [appState.meta.assets, editorRef, lastHydratedPageIdRef, page, selectionRangeRef])

  useEffect(() => {
    noteCanvasScrollRef.current?.scrollTo({ top: 0, behavior: 'auto' })
  }, [noteCanvasScrollRef, page?.id])

  useEffect(() => {
    const captureSelection = () => {
      const selection = window.getSelection()
      if (!selection || selection.rangeCount === 0) return
      const range = selection.getRangeAt(0)
      if (editorRef.current?.contains(range.commonAncestorContainer)) {
        selectionRangeRef.current = range.cloneRange()
        if (selection.isCollapsed) {
          keepCaretInView()
        }
      }
    }

    document.addEventListener('selectionchange', captureSelection)
    return () => {
      document.removeEventListener('selectionchange', captureSelection)
    }
  }, [editorRef, keepCaretInView, selectionRangeRef])

  useEffect(() => {
    const handleWindowCopy = (event: ClipboardEvent) => {
      const target = event.target
      if (!(target instanceof Node)) return
      const root = document.querySelector('.onenote-window')
      if (!root?.contains(target)) return

      const selection = window.getSelection()?.toString() ?? ''
      const activeElement = document.activeElement
      const inputSelection =
        activeElement instanceof HTMLInputElement || activeElement instanceof HTMLTextAreaElement
          ? activeElement.value.slice(activeElement.selectionStart ?? 0, activeElement.selectionEnd ?? 0)
          : ''

      const rawText = inputSelection || selection
      if (!rawText.trim()) return

      event.preventDefault()
      event.clipboardData?.setData('text/plain', normalizeTerminalText(rawText))
    }

    const handleClickOutside = (event: MouseEvent) => {
      if (
        styleMenuRef.current &&
        !styleMenuRef.current.contains(event.target as Node) &&
        !styleMenuPanelRef.current?.contains(event.target as Node)
      ) {
        setIsStyleMenuOpen(false)
      }
      if (
        fontMenuRef.current &&
        !fontMenuRef.current.contains(event.target as Node) &&
        !fontMenuPanelRef.current?.contains(event.target as Node)
      ) {
        setIsFontMenuOpen(false)
      }
      if (
        fontSizeMenuRef.current &&
        !fontSizeMenuRef.current.contains(event.target as Node) &&
        !fontSizeMenuPanelRef.current?.contains(event.target as Node)
      ) {
        setIsFontSizeMenuOpen(false)
      }
      if (
        titlebarSearchRef.current &&
        !titlebarSearchRef.current.contains(event.target as Node) &&
        !searchResultsPanelRef.current?.contains(event.target as Node)
      ) {
        setQuery('')
      }
    }

    window.addEventListener('copy', handleWindowCopy)
    window.addEventListener('mousedown', handleClickOutside)
    return () => {
      window.removeEventListener('copy', handleWindowCopy)
      window.removeEventListener('mousedown', handleClickOutside)
    }
  }, [
    fontMenuPanelRef,
    fontMenuRef,
    fontSizeMenuPanelRef,
    fontSizeMenuRef,
    searchResultsPanelRef,
    setIsFontMenuOpen,
    setIsFontSizeMenuOpen,
    setIsStyleMenuOpen,
    setQuery,
    styleMenuPanelRef,
    styleMenuRef,
    titlebarSearchRef,
  ])

  const syncEditorContent = useCallback(() => {
    updatePage({
      content: dehydratePageContent(sanitizePastedHtml(editorRef.current?.innerHTML ?? page?.content ?? '')),
    })
  }, [editorRef, page?.content, updatePage])

  return {
    keepCaretInView,
    pauseAutoFollow,
    syncEditorContent,
  }
}
