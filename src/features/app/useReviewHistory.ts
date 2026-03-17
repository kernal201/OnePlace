import { useEffect, useMemo, type Dispatch, type RefObject, type SetStateAction } from 'react'
import {
  buildSnippet,
  countTextMatchesInHtml,
  extractSnippetText,
  flattenPages,
  hydratePageContent,
  replaceTextInHtml,
} from '../../app/appModel'
import type { AppState, Page, ReviewScope, Notebook, Section, SectionGroup } from '../../app/appModel'

type UseReviewHistoryArgs = {
  appState: AppState
  editorRef: RefObject<HTMLDivElement | null>
  notebook: Notebook | undefined
  page: Page | undefined
  reviewFind: string
  reviewReplace: string
  reviewScope: ReviewScope
  reviewScopeLabels: Record<ReviewScope, string>
  section: Section | undefined
  sectionGroup: SectionGroup | undefined
  selectedHistoryVersionId: string
  selectedTemplateId: string
  pageTemplates: Array<{ html: string; id: string; label: string }>
  setAppState: Dispatch<SetStateAction<AppState>>
  setSaveLabel: (value: string) => void
  setSelectedHistoryVersionId: Dispatch<SetStateAction<string>>
}

export const useReviewHistory = ({
  appState,
  editorRef,
  notebook,
  page,
  reviewFind,
  reviewReplace,
  reviewScope,
  reviewScopeLabels,
  section,
  sectionGroup,
  selectedHistoryVersionId,
  selectedTemplateId,
  pageTemplates,
  setAppState,
  setSaveLabel,
  setSelectedHistoryVersionId,
}: UseReviewHistoryArgs) => {
  const reviewScopePages = useMemo(() => {
    if (reviewScope === 'page') {
      return page ? [page] : []
    }

    const notebookTargets =
      reviewScope === 'section'
        ? notebook && sectionGroup && section
          ? [notebook]
          : []
        : reviewScope === 'notebook' && notebook
          ? [notebook]
          : appState.notebooks

    return notebookTargets.flatMap((entry) =>
      entry.sectionGroups.flatMap((group) =>
        group.sections.flatMap((part) => {
          if (
            reviewScope === 'section' &&
            (entry.id !== notebook?.id || group.id !== sectionGroup?.id || part.id !== section?.id)
          ) {
            return []
          }

          return flattenPages(part.pages, 0, true).map((item) => item.page)
        }),
      ),
    )
  }, [appState.notebooks, notebook, page, reviewScope, section, sectionGroup])

  const reviewMatchCount = useMemo(
    () =>
      reviewScopePages.reduce(
        (total, targetPage) => total + countTextMatchesInHtml(targetPage.content ?? '', reviewFind),
        0,
      ),
    [reviewFind, reviewScopePages],
  )

  const currentPageVersions = useMemo(() => (page ? appState.meta.pageVersions[page.id] ?? [] : []), [appState.meta.pageVersions, page])
  const selectedHistoryVersion = useMemo(
    () => currentPageVersions.find((version) => version.id === selectedHistoryVersionId) ?? currentPageVersions[0] ?? null,
    [currentPageVersions, selectedHistoryVersionId],
  )
  const selectedTemplate = useMemo(
    () => pageTemplates.find((entry) => entry.id === selectedTemplateId) ?? pageTemplates[0] ?? null,
    [pageTemplates, selectedTemplateId],
  )
  const historyPreviewText = selectedHistoryVersion ? extractSnippetText(selectedHistoryVersion.content).slice(0, 180) : ''

  useEffect(() => {
    if (!page) {
      setSelectedHistoryVersionId('')
      return
    }

    const latestVersionId = currentPageVersions[0]?.id ?? ''
    setSelectedHistoryVersionId((current) =>
      current && currentPageVersions.some((version) => version.id === current) ? current : latestVersionId,
    )
  }, [currentPageVersions, page, setSelectedHistoryVersionId])

  const replaceInReviewScope = () => {
    if (!reviewFind.trim() || reviewScopePages.length === 0) return

    const targetIds = new Set(reviewScopePages.map((targetPage) => targetPage.id))
    let replacements = 0

    setAppState((current) => ({
      ...current,
      notebooks: current.notebooks.map((entry) => ({
        ...entry,
        sectionGroups: entry.sectionGroups.map((group) => ({
          ...group,
          sections: group.sections.map((part) => ({
            ...part,
            pages: part.pages.map(function transformPageTree(note): Page {
              const nextChildren = note.children.map(transformPageTree)
              if (!targetIds.has(note.id)) {
                return nextChildren === note.children ? note : { ...note, children: nextChildren }
              }
              const nextContent = replaceTextInHtml(note.content, reviewFind, reviewReplace)
              replacements += countTextMatchesInHtml(note.content, reviewFind)
              if (nextContent === note.content && nextChildren === note.children) return note
              return {
                ...note,
                children: nextChildren,
                content: nextContent,
                snippet: buildSnippet(note.title, nextContent),
                updatedAt: new Date().toISOString(),
              }
            }),
          })),
        })),
      })),
    }))

    if (page && targetIds.has(page.id) && editorRef.current) {
      const nextCurrentContent = replaceTextInHtml(page.content, reviewFind, reviewReplace)
      editorRef.current.innerHTML = hydratePageContent(nextCurrentContent, appState.meta.assets)
    }

    setSaveLabel(
      replacements > 0
        ? `Replaced ${reviewFind} across ${reviewScopeLabels[reviewScope].toLowerCase()}`
        : `No matches to replace in ${reviewScopeLabels[reviewScope].toLowerCase()}`,
    )
  }

  return {
    currentPageVersions,
    historyPreviewText,
    replaceInReviewScope,
    reviewMatchCount,
    selectedHistoryVersion,
    selectedTemplate,
  }
}
