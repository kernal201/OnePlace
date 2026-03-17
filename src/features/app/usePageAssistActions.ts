import type { MouseEvent as ReactMouseEvent, Dispatch, SetStateAction } from 'react'
import { createId, extractSnippetText, hydratePageContent } from '../../app/appModel'
import type { AppAsset, CopilotMessage, Page } from '../../app/appModel'
import { buildCopilotResponse } from './copilot'

type Args = {
  appAssets: Record<string, AppAsset>
  copilotDraft: string
  handleEditorAssetClick: (target: HTMLElement) => boolean
  insertHtmlAtSelection: (html: string) => void
  keepCaretInView: (behavior?: ScrollBehavior) => void
  openLinkedPage: (pageId: string) => void
  page: Page | undefined
  setCopilotDraft: (value: string) => void
  setCopilotMessages: Dispatch<SetStateAction<CopilotMessage[]>>
  setEditorZoom: Dispatch<SetStateAction<number>>
  setSaveLabel: (value: string) => void
  syncEditorContent: () => void
}

export const usePageAssistActions = ({
  appAssets,
  copilotDraft,
  handleEditorAssetClick,
  insertHtmlAtSelection,
  keepCaretInView,
  openLinkedPage,
  page,
  setCopilotDraft,
  setCopilotMessages,
  setEditorZoom,
  setSaveLabel,
  syncEditorContent,
}: Args) => {
  const runCopilotPrompt = (prompt: string) => {
    if (!page) return
    const noteText = extractSnippetText(hydratePageContent(page.content, appAssets))
      .replace(/\s+/g, ' ')
      .trim()
    const response = buildCopilotResponse(prompt, noteText)
    insertHtmlAtSelection(response)
    setCopilotMessages((current) => [{ id: createId(), prompt, response }, ...current].slice(0, 6))
    setSaveLabel('Copilot inserted content into the page')
  }

  const submitCopilotDraft = () => {
    const prompt = copilotDraft.trim()
    if (!prompt) return
    runCopilotPrompt(prompt)
    setCopilotDraft('')
  }

  const adjustEditorZoom = (delta: number) => {
    setEditorZoom((current) => Math.min(1.6, Math.max(0.7, Math.round((current + delta) * 100) / 100)))
  }

  const handleEditorClick = (event: ReactMouseEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement
    const linkedPageId = target.closest<HTMLElement>('[data-page-id]')?.dataset.pageId
    if (linkedPageId) {
      event.preventDefault()
      openLinkedPage(linkedPageId)
      return
    }

    if (handleEditorAssetClick(target)) {
      return
    }

    window.setTimeout(() => {
      syncEditorContent()
      keepCaretInView()
    }, 0)
  }

  return {
    adjustEditorZoom,
    handleEditorClick,
    runCopilotPrompt,
    submitCopilotDraft,
  }
}
