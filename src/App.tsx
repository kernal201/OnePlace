import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent as ReactChangeEvent,
  type ClipboardEvent as ReactClipboardEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import { createPortal } from 'react-dom'
import {
  AlignLeftIcon,
  AttachmentIcon,
  BoldIcon,
  BrushIcon,
  BulletsIcon,
  ChevronDownIcon,
  CopyIcon,
  CutIcon,
  DeleteIcon,
  EditIcon,
  FolderIcon,
  FormatMotivationIcon,
  ImageIcon,
  IndentIcon,
  InsertFormattingIcon,
  ItalicIcon,
  LinkIcon,
  ListLinesIcon,
  NotebookStackIcon,
  OneNoteLogoIcon,
  PasteIcon,
  PenIcon,
  PersonIcon,
  ProjectIcon,
  SaveIcon,
  SearchIcon,
  SectionBookIcon,
  SettingsIcon,
  ShowIcon,
  SortIcon,
  SubpageIcon,
  TableIcon,
  TagsIcon,
  TextSizeDownIcon,
  TextSizeUpIcon,
  UnderlineIcon,
  UndoIcon,
} from './components/Icons'
import {
  checkForDesktopUpdate,
  downloadAndInstallDesktopUpdate,
  exportNotebookDirectory,
  getDesktopAppInfo,
  loadDesktopData,
  openNotebookDirectory,
  pickNotebookDirectory,
  saveDesktopData,
} from './lib/desktop'
import './App.css'

type Page = {
  accent: string
  children: Page[]
  content: string
  createdAt: string
  id: string
  inkStrokes: InkStroke[]
  isCollapsed: boolean
  snippet: string
  tags: string[]
  task: PageTask | null
  title: string
  updatedAt: string
}

type Section = {
  color: string
  id: string
  passwordHash: string | null
  passwordHint: string
  name: string
  pages: Page[]
}

type InkPoint = {
  x: number
  y: number
}

type InkStroke = {
  color: string
  id: string
  points: InkPoint[]
  width: number
}

type PageTask = {
  dueAt: string | null
  status: 'done' | 'open'
}

type SectionGroup = {
  id: string
  isCollapsed: boolean
  name: string
  sections: Section[]
}

type Notebook = {
  color: string
  icon: string
  id: string
  name: string
  sectionGroups: SectionGroup[]
}

type SearchScope = 'section' | 'notebook' | 'all'

type PageSortMode = 'manual' | 'updated-desc' | 'updated-asc' | 'title-asc' | 'title-desc' | 'created-desc'

type AppAsset = {
  createdAt: string
  dataUrl: string
  id: string
  kind: 'audio' | 'file' | 'image' | 'printout'
  mimeType: string
  name: string
  sizeLabel: string
}

type PageVersion = {
  content: string
  id: string
  savedAt: string
  title: string
}

type AppMeta = {
  assets: Record<string, AppAsset>
  pageSortMode: PageSortMode
  pageVersions: Record<string, PageVersion[]>
  recentPageIds: string[]
  searchScope: SearchScope
}

type ContextMenuItem = {
  danger?: boolean
  disabled?: boolean
  label: string
  onSelect: () => void
}

type ContextMenuState = {
  items: ContextMenuItem[]
  x: number
  y: number
}

type AppState = {
  meta: AppMeta
  notebooks: Notebook[]
  selectedNotebookId: string
  selectedPageId: string
  selectedSectionGroupId: string
  selectedSectionId: string
}

type SearchResult = {
  groupId: string
  groupName: string
  isSubpage: boolean
  notebookId: string
  notebookName: string
  page: Page
  sectionId: string
  sectionName: string
}

type VisiblePage = {
  depth: number
  page: Page
}

type PageLocation = {
  depth: number
  index: number
  page: Page
  parentId?: string
}

type DragState =
  | { type: 'notebook'; notebookId: string }
  | { type: 'section-group'; groupId: string }
  | { type: 'section'; groupId: string; sectionId: string }
  | { type: 'page'; pageId: string }

type DropTarget =
  | { type: 'notebook'; notebookId: string; position: 'before' | 'after' }
  | { type: 'section-group'; groupId: string; position: 'before' | 'after' }
  | { type: 'section'; groupId: string; position: 'inside' }
  | { type: 'section'; groupId: string; sectionId: string; position: 'before' | 'after' }
  | { type: 'page'; pageId: string; position: 'before' | 'after' }

type DragPosition = {
  x: number
  y: number
}

type RecentNotebookEntry = {
  name: string
  path: string
}

const ribbonTabs = ['File', 'Home', 'Insert', 'Draw', 'History', 'Review', 'View']
type RibbonTab = (typeof ribbonTabs)[number]
const RECENT_NOTEBOOKS_KEY = 'oneplace-recent-notebooks'
const LAST_OPENED_NOTEBOOK_KEY = 'oneplace-last-opened-notebook'

const stylePresets = [
  {
    id: 'heading-1',
    label: 'Heading 1',
    html: '<h1>Heading 1</h1>',
  },
  {
    id: 'heading-2',
    label: 'Heading 2',
    html: '<h2>Heading 2</h2>',
  },
  {
    id: 'callout',
    label: 'Callout',
    html: '<blockquote>Callout</blockquote>',
  },
  {
    id: 'code-block',
    label: 'Code Block',
    html: '<pre><code>terminal-safe-text</code></pre>',
  },
]

const fontFamilies = ['Calibri', 'Segoe UI', 'Arial', 'Georgia', 'Times New Roman', 'Consolas']
const fontSizes = [
  { command: '1', label: '8' },
  { command: '2', label: '10' },
  { command: '3', label: '11' },
  { command: '4', label: '14' },
  { command: '5', label: '18' },
  { command: '6', label: '24' },
]

const pageTemplates = [
  {
    id: 'meeting-notes',
    label: 'Meeting Notes',
    html: `
      <section class="template-block">
        <h2>Meeting Notes</h2>
        <p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
        <p><strong>Attendees:</strong> </p>
        <h3>Agenda</h3>
        <ul><li></li></ul>
        <h3>Decisions</h3>
        <ul><li></li></ul>
        <h3>Action Items</h3>
        <ul class="checklist"><li><label><input type="checkbox" /> </label></li></ul>
      </section>
    `,
  },
  {
    id: 'project-brief',
    label: 'Project Brief',
    html: `
      <section class="template-block">
        <h2>Project Brief</h2>
        <p><strong>Objective:</strong> </p>
        <p><strong>Owner:</strong> </p>
        <h3>Scope</h3>
        <ul><li></li></ul>
        <h3>Risks</h3>
        <ul><li></li></ul>
        <h3>Timeline</h3>
        <table class="action-table">
          <tbody>
            <tr><th>Milestone</th><th>Date</th><th>Status</th></tr>
            <tr><td></td><td></td><td></td></tr>
          </tbody>
        </table>
      </section>
    `,
  },
]

const defaultTask = (): PageTask => ({
  dueAt: null,
  status: 'open',
})

const defaultAppMeta = (): AppMeta => ({
  assets: {},
  pageSortMode: 'manual',
  pageVersions: {},
  recentPageIds: [],
  searchScope: 'all',
})

const pageSortModeLabels: Record<PageSortMode, string> = {
  manual: 'Manual',
  'updated-desc': 'Newest edited',
  'updated-asc': 'Oldest edited',
  'title-asc': 'Title A-Z',
  'title-desc': 'Title Z-A',
  'created-desc': 'Newest created',
}

const searchScopeLabels: Record<SearchScope, string> = {
  all: 'All notebooks',
  notebook: 'Current notebook',
  section: 'Current section',
}

const formatDate = (value: string) =>
  new Intl.DateTimeFormat('en-CA', {
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    month: 'short',
  }).format(new Date(value))

const formatPageDate = (value: string) =>
  new Intl.DateTimeFormat('en-CA', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  }).format(new Date(value))

const formatPageTime = (value: string) =>
  new Intl.DateTimeFormat('en-CA', {
    hour: 'numeric',
    minute: '2-digit',
  }).format(new Date(value))

const createId = () => crypto.randomUUID()

const createPage = (
  title: string,
  snippet: string,
  accent: string,
  content: string,
  children: Page[] = [],
): Page => {
  const now = new Date().toISOString()
  return {
    accent,
    children,
    content,
    createdAt: now,
    id: createId(),
    inkStrokes: [],
    isCollapsed: false,
    snippet: snippet || buildSnippet(title, content, now),
    tags: [],
    task: null,
    title,
    updatedAt: now,
  }
}

const extractSnippetText = (content: string) => {
  if (typeof DOMParser === 'undefined') {
    return content.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()
  }

  const doc = new DOMParser().parseFromString(content, 'text/html')

  for (const note of doc.querySelectorAll('.audio-note')) {
    note.textContent = 'Audio note'
  }

  for (const card of doc.querySelectorAll('.attachment-card')) {
    const title = card.querySelector('.attachment-title')?.textContent?.trim()
    const meta = card.querySelector('.attachment-meta')?.textContent?.trim()
    card.textContent = [title, meta].filter(Boolean).join(' ')
  }

  for (const card of doc.querySelectorAll('.printout-card')) {
    const caption = card.querySelector('.printout-caption')?.textContent?.trim()
    card.textContent = caption || 'Printout'
  }

  const plain = doc.body.textContent?.replace(/\s+/g, ' ').trim() ?? ''
  return plain.replace(/&nbsp;|&#160;/gi, ' ').replace(/\s+/g, ' ').trim()
}

const buildSnippet = (title: string, content: string, timestamp = new Date().toISOString()) => {
  const plain = extractSnippetText(content)
  return `${formatPageDate(timestamp)}\n${plain || title}`.slice(0, 120)
}

const escapeHtml = (value: string) =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')

const escapeAttribute = (value: string) => escapeHtml(value)

const linkifyPlainText = (value: string) => {
  const expression = /(https?:\/\/[^\s<]+|www\.[^\s<]+)/gi
  let cursor = 0
  let html = ''

  for (const match of value.matchAll(expression)) {
    const rawUrl = match[0]
    const index = match.index ?? 0
    html += escapeHtml(value.slice(cursor, index))
    const href = rawUrl.startsWith('http') ? rawUrl : `https://${rawUrl}`
    html += `<a href="${escapeAttribute(href)}" target="_blank" rel="noreferrer">${escapeHtml(rawUrl)}</a>`
    cursor = index + rawUrl.length
  }

  html += escapeHtml(value.slice(cursor))
  return html
}

const plainTextToHtml = (value: string) =>
  value
    .split(/\n{2,}/)
    .map((block) => `<p>${linkifyPlainText(block.trim()).replace(/\n/g, '<br />')}</p>`)
    .join('')

const getFloatingMenuStyle = (element: HTMLDivElement | null, width?: number) => {
  if (!element) return undefined
  const rect = element.getBoundingClientRect()

  return {
    left: rect.left,
    minWidth: width ?? rect.width,
    position: 'fixed' as const,
    top: rect.bottom + 2,
  }
}

const getFloatingPanelStyle = (element: HTMLDivElement | null) => {
  if (!element) return undefined
  const rect = element.getBoundingClientRect()

  return {
    left: rect.left,
    position: 'fixed' as const,
    top: rect.bottom + 2,
    width: rect.width,
  }
}

const sanitizePastedHtml = (value: string) => {
  const parser = new DOMParser()
  const doc = parser.parseFromString(value, 'text/html')
  const allowedTags = new Set([
    'a',
    'audio',
    'blockquote',
    'br',
    'code',
    'div',
    'em',
    'figcaption',
    'figure',
    'font',
    'h1',
    'h2',
    'h3',
    'hr',
    'iframe',
    'img',
    'input',
    'label',
    'li',
    'ol',
    'p',
    'pre',
    'section',
    'span',
    'strong',
    'table',
    'tbody',
    'td',
    'th',
    'thead',
    'tr',
    'u',
    'ul',
  ])

  const unwrapNode = (node: Element) => {
    const parent = node.parentNode
    if (!parent) return
    while (node.firstChild) {
      parent.insertBefore(node.firstChild, node)
    }
    parent.removeChild(node)
  }

  const scrubNode = (node: Element) => {
    const tag = node.tagName.toLowerCase()
    if (!allowedTags.has(tag)) {
      unwrapNode(node)
      return
    }

    for (const attribute of [...node.attributes]) {
      const attrName = attribute.name.toLowerCase()
      const attrValue = attribute.value

      if (tag === 'a' && attrName === 'href') {
        if (!/^https?:|^mailto:|^#page:/.test(attrValue)) {
          node.removeAttribute(attribute.name)
        } else if (/^https?:|^mailto:/.test(attrValue)) {
          node.setAttribute('target', '_blank')
          node.setAttribute('rel', 'noreferrer')
        }
        continue
      }

      if (tag === 'img' && attrName === 'src') {
        if (!/^data:image\/|^https?:/.test(attrValue)) {
          node.removeAttribute(attribute.name)
        }
        continue
      }

      if (tag === 'audio' && attrName === 'src') {
        if (!/^data:audio\/|^https?:/.test(attrValue)) {
          node.removeAttribute(attribute.name)
        }
        continue
      }

      if (tag === 'audio' && attrName === 'controls') {
        continue
      }

      if (tag === 'iframe' && attrName === 'src') {
        if (!/^data:application\/pdf(?:;base64)?,/i.test(attrValue)) {
          node.removeAttribute(attribute.name)
        }
        continue
      }

      if (tag === 'iframe' && attrName === 'title') {
        continue
      }

      if (tag === 'input' && (attrName === 'type' || attrName === 'checked')) {
        if (attrName === 'type' && attrValue !== 'checkbox') {
          node.removeAttribute(attribute.name)
        }
        continue
      }

      if (tag === 'font' && (attrName === 'face' || attrName === 'size' || attrName === 'color')) {
        continue
      }

      if (attrName === 'class') {
        const allowedClasses = attrValue
          .split(/\s+/)
          .filter((className) =>
            [
              'action-table',
              'audio-note',
              'attachment-card',
              'attachment-meta',
              'attachment-title',
              'checklist',
              'embedded-image',
              'internal-page-link',
              'page-template-card',
              'printout-caption',
              'printout-card',
              'printout-preview',
              'printout-preview-shell',
              'template-block',
            ].includes(className),
          )
        if (allowedClasses.length > 0) {
          node.setAttribute('class', allowedClasses.join(' '))
        } else {
          node.removeAttribute('class')
        }
        continue
      }

      if (
        attrName === 'contenteditable' ||
        attrName === 'data-asset-id' ||
        attrName === 'data-page-id' ||
        attrName === 'data-file-name' ||
        attrName === 'data-download-url'
      ) {
        continue
      }

      node.removeAttribute(attribute.name)
    }

    if (tag === 'span' && !node.attributes.length) {
      const text = node.textContent ?? ''
      node.replaceWith(doc.createTextNode(text))
      return
    }

    for (const child of [...node.children]) {
      scrubNode(child)
    }
  }

  for (const child of [...doc.body.children]) {
    scrubNode(child)
  }

  return doc.body.innerHTML
}

const normalizeTerminalText = (value: string) =>
  value
    .normalize('NFKD')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\u2026/g, '...')
    .replace(/\u00A0/g, ' ')
    .replace(/\u2022/g, '*')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[^\t\n\r -~]/g, '')

const hydratePageContent = (value: string, assets: Record<string, AppAsset>) => {
  const parser = new DOMParser()
  const doc = parser.parseFromString(value, 'text/html')

  for (const image of doc.querySelectorAll<HTMLElement>('[data-asset-id]')) {
    const assetId = image.dataset.assetId
    if (!assetId) continue
    const asset = assets[assetId]
    if (!asset) continue

    if (image instanceof HTMLImageElement) {
      image.src = asset.dataUrl
      continue
    }

    if (image instanceof HTMLAudioElement) {
      image.src = asset.dataUrl
      continue
    }

    if (image.classList.contains('attachment-card') || image.classList.contains('printout-card')) {
      image.dataset.downloadUrl = asset.dataUrl
      image.dataset.fileName = asset.name
    }
  }

  for (const frame of doc.querySelectorAll<HTMLIFrameElement>('iframe[data-asset-id]')) {
    const assetId = frame.dataset.assetId
    if (!assetId) continue
    const asset = assets[assetId]
    if (!asset) continue
    frame.src = asset.dataUrl
  }

  return doc.body.innerHTML
}

const dehydratePageContent = (value: string) => {
  const parser = new DOMParser()
  const doc = parser.parseFromString(value, 'text/html')

  for (const image of doc.querySelectorAll<HTMLElement>('[data-asset-id]')) {
    if (image instanceof HTMLImageElement) {
      image.removeAttribute('src')
      continue
    }

    if (image instanceof HTMLAudioElement) {
      image.removeAttribute('src')
      continue
    }

    if (image.classList.contains('attachment-card') || image.classList.contains('printout-card')) {
      image.removeAttribute('data-download-url')
      image.removeAttribute('data-file-name')
    }
  }

  for (const frame of doc.querySelectorAll<HTMLIFrameElement>('iframe[data-asset-id]')) {
    frame.removeAttribute('src')
  }

  return doc.body.innerHTML
}

const recordPageVersion = (
  versions: Record<string, PageVersion[]>,
  pageId: string,
  title: string,
  content: string,
) => {
  const existing = versions[pageId] ?? []
  const latest = existing[0]
  if (latest && latest.title === title && latest.content === content) {
    return versions
  }

  return {
    ...versions,
    [pageId]: [
      {
        content,
        id: createId(),
        savedAt: new Date().toISOString(),
        title,
      },
      ...existing,
    ].slice(0, 20),
  }
}

const hashSecret = async (value: string) => {
  const encoded = new TextEncoder().encode(value)
  const digest = await crypto.subtle.digest('SHA-256', encoded)
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}

const sortPagesTree = (pages: Page[], mode: PageSortMode): Page[] => {
  if (mode === 'manual') return pages

  const comparePages = (left: Page, right: Page) => {
    switch (mode) {
      case 'updated-desc':
        return new Date(right.updatedAt).getTime() - new Date(left.updatedAt).getTime()
      case 'updated-asc':
        return new Date(left.updatedAt).getTime() - new Date(right.updatedAt).getTime()
      case 'title-asc':
        return left.title.localeCompare(right.title)
      case 'title-desc':
        return right.title.localeCompare(left.title)
      case 'created-desc':
        return new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime()
      default:
        return 0
    }
  }

  return [...pages]
    .sort(comparePages)
    .map((page) => ({
      ...page,
      children: sortPagesTree(page.children, mode),
    }))
}

const createSectionGroup = (name: string, sections: Section[] = []): SectionGroup => ({
  id: createId(),
  isCollapsed: false,
  name,
  sections,
})

const flattenPages = (pages: Page[], depth = 0, includeCollapsedChildren = false): VisiblePage[] =>
  pages.flatMap((page) => [
    { depth, page },
    ...(
      includeCollapsedChildren || !page.isCollapsed
        ? flattenPages(page.children, depth + 1, includeCollapsedChildren)
        : []
    ),
  ])

const findPageById = (pages: Page[], pageId: string): Page | undefined => {
  for (const page of pages) {
    if (page.id === pageId) return page
    const childMatch = findPageById(page.children, pageId)
    if (childMatch) return childMatch
  }
  return undefined
}

const updateNestedPages = (pages: Page[], pageId: string, updater: (page: Page) => Page): Page[] =>
  pages.map((page) =>
    page.id === pageId
      ? updater(page)
      : { ...page, children: updateNestedPages(page.children, pageId, updater) },
  )

const reorderItems = <T extends { id: string }>(
  items: T[],
  draggedId: string,
  targetId: string,
  position: 'before' | 'after' = 'after',
): T[] => {
  const draggedIndex = items.findIndex((item) => item.id === draggedId)
  const targetIndex = items.findIndex((item) => item.id === targetId)
  if (draggedIndex === -1 || targetIndex === -1 || draggedIndex === targetIndex) {
    return items
  }

  const draggedItem = items[draggedIndex]
  const nextItems = items.filter((item) => item.id !== draggedId)
  const resolvedTargetIndex = nextItems.findIndex((item) => item.id === targetId)
  const insertIndex = resolvedTargetIndex + (position === 'after' ? 1 : 0)
  nextItems.splice(insertIndex, 0, draggedItem)
  return nextItems
}

const removePageById = (pages: Page[], pageId: string): { page?: Page; pages: Page[] } => {
  const nextPages: Page[] = []

  for (const page of pages) {
    if (page.id === pageId) {
      return { page, pages: [...nextPages, ...pages.slice(nextPages.length + 1)] }
    }

    const nested = removePageById(page.children, pageId)
    if (nested.page) {
      nextPages.push({ ...page, children: nested.pages })
      return { page: nested.page, pages: [...nextPages, ...pages.slice(nextPages.length)] }
    }

    nextPages.push(page)
  }

  return { pages }
}

const insertPageRelative = (
  pages: Page[],
  targetId: string,
  pageToInsert: Page,
  position: 'before' | 'after',
): Page[] =>
  pages.flatMap((page) => {
    if (page.id === targetId) {
      return position === 'before' ? [pageToInsert, page] : [page, pageToInsert]
    }

    if (findPageById(page.children, targetId)) {
      return [{ ...page, children: insertPageRelative(page.children, targetId, pageToInsert, position) }]
    }

    return [page]
  })

const pageContainsId = (page: Page, targetId: string): boolean =>
  page.id === targetId || page.children.some((child) => pageContainsId(child, targetId))

const hasChildPageSelected = (page: Page, selectedPageId: string) =>
  page.children.some((child) => pageContainsId(child, selectedPageId))

const getDropPosition = (event: ReactPointerEvent<HTMLElement>): 'before' | 'after' => {
  const bounds = event.currentTarget.getBoundingClientRect()
  return event.clientY < bounds.top + bounds.height / 2 ? 'before' : 'after'
}

const countPages = (pages: Page[]) => flattenPages(pages, 0, true).length

const findPageLocation = (
  pages: Page[],
  pageId: string,
  depth = 0,
  parentId?: string,
): PageLocation | undefined => {
  for (let index = 0; index < pages.length; index += 1) {
    const page = pages[index]
    if (page.id === pageId) {
      return { depth, index, page, parentId }
    }

    const nested = findPageLocation(page.children, pageId, depth + 1, page.id)
    if (nested) return nested
  }

  return undefined
}

const promotePageOneLevel = (pages: Page[], pageId: string): { moved: boolean; pages: Page[] } => {
  for (let index = 0; index < pages.length; index += 1) {
    const page = pages[index]
    const childIndex = page.children.findIndex((child) => child.id === pageId)

    if (childIndex !== -1) {
      const promotedPage = page.children[childIndex]
      const nextSiblings = [...pages]
      nextSiblings[index] = {
        ...page,
        children: page.children.filter((_, candidateIndex) => candidateIndex !== childIndex),
      }
      nextSiblings.splice(index + 1, 0, promotedPage)
      return { moved: true, pages: nextSiblings }
    }

    const nested = promotePageOneLevel(page.children, pageId)
    if (nested.moved) {
      const nextSiblings = [...pages]
      nextSiblings[index] = { ...page, isCollapsed: false, children: nested.pages }
      return { moved: true, pages: nextSiblings }
    }
  }

  return { moved: false, pages }
}

const demotePageOneLevel = (pages: Page[], pageId: string): { moved: boolean; pages: Page[] } => {
  for (let index = 0; index < pages.length; index += 1) {
    const page = pages[index]

    if (page.id === pageId) {
      if (index === 0) return { moved: false, pages }

      const nextSiblings = [...pages]
      const [movedPage] = nextSiblings.splice(index, 1)
      const previousPage = nextSiblings[index - 1]
      nextSiblings[index - 1] = {
        ...previousPage,
        isCollapsed: false,
        children: [...previousPage.children, movedPage],
      }
      return { moved: true, pages: nextSiblings }
    }

    const nested = demotePageOneLevel(page.children, pageId)
    if (nested.moved) {
      const nextSiblings = [...pages]
      nextSiblings[index] = { ...page, isCollapsed: false, children: nested.pages }
      return { moved: true, pages: nextSiblings }
    }
  }

  return { moved: false, pages }
}

const accentPalette = ['#3784d6', '#3c9fa7', '#d1629b', '#d77f9c', '#4c75b8', '#b15fab']

const createStarterState = (): AppState => {
  const createStarterPage = (title: string, accent: string) =>
    createPage(title, '', accent, `<p>${title} notes.</p>`)

  const createStarterSection = (name: string, color: string, pageTitles?: string[]): Section => ({
    color,
    id: createId(),
    name,
    pages: (pageTitles?.length ? pageTitles : ['Untitled Page']).map((title) => createStarterPage(title, color)),
    passwordHash: null,
    passwordHint: '',
  })

  const workNotebook: Notebook = {
    color: '#7e42b3',
    icon: 'book',
    id: createId(),
    name: 'Work Notebook',
    sectionGroups: [
      createSectionGroup('Sections', [
        createStarterSection('Administration', '#4c75b8'),
        createStarterSection('Meetings', '#3c9fa7'),
        createStarterSection('Product Ideas', '#d1629b'),
        createStarterSection('Onboarding', '#d77f9c'),
        createStarterSection('Email List', '#4c75b8'),
        createStarterSection('Customers', '#b15fab'),
        createStarterSection('Schedules', '#3784d6'),
        createStarterSection('Resources', '#4c75b8'),
        createStarterSection('Inventory', '#3c9fa7', [
          'Spokes',
          'Cogs',
          'Saddle Inventory',
          'Tires',
          'Pedals',
          'Rims',
          'Wheels',
          'Down Tubes',
        ]),
        createStarterSection('The Team', '#8b74bb'),
      ]),
    ],
  }

  const travelNotebook: Notebook = {
    color: '#b33f66',
    icon: 'folder',
    id: createId(),
    name: 'Travel Journal',
    sectionGroups: [
      createSectionGroup('Sections', [
        createStarterSection('Wishlist', '#b33f66'),
        createStarterSection('Italy', '#d8763c'),
        createStarterSection('Iceland', '#6387c7'),
      ]),
    ],
  }

  const defaultGroup = workNotebook.sectionGroups[0]
  const defaultSection = defaultGroup.sections[0]
  const defaultPage = defaultSection.pages[0]
  return {
    meta: {
      ...defaultAppMeta(),
      recentPageIds: [defaultPage.id],
    },
    notebooks: [workNotebook, travelNotebook],
    selectedNotebookId: workNotebook.id,
    selectedSectionGroupId: defaultGroup.id,
    selectedSectionId: defaultSection.id,
    selectedPageId: defaultPage.id,
  }
}

const loadRecentNotebookEntries = (): RecentNotebookEntry[] => {
  try {
    const raw = localStorage.getItem(RECENT_NOTEBOOKS_KEY)
    if (!raw) return []
    const parsed = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed.filter(
      (item): item is RecentNotebookEntry =>
        Boolean(
          item &&
            typeof item === 'object' &&
            'name' in item &&
            typeof item.name === 'string' &&
            'path' in item &&
            typeof item.path === 'string',
        ),
    )
  } catch {
    return []
  }
}

const saveRecentNotebookEntries = (entries: RecentNotebookEntry[]) => {
  localStorage.setItem(RECENT_NOTEBOOKS_KEY, JSON.stringify(entries.slice(0, 8)))
}

const loadLastOpenedNotebookPath = () => localStorage.getItem(LAST_OPENED_NOTEBOOK_KEY)

const saveLastOpenedNotebookPath = (path: string) => {
  localStorage.setItem(LAST_OPENED_NOTEBOOK_KEY, path)
}

const mergeNotebookIntoState = (current: AppState, openedNotebook: Notebook): AppState =>
  ensureSelection({
    ...current,
    notebooks: current.notebooks.some((item) => item.id === openedNotebook.id)
      ? current.notebooks.map((item) => (item.id === openedNotebook.id ? openedNotebook : item))
      : [...current.notebooks, openedNotebook],
    selectedNotebookId: openedNotebook.id,
  })

const normalizeAppState = (input: unknown): AppState => {
  if (!input || typeof input !== 'object' || !('notebooks' in input)) {
    return createStarterState()
  }

  const rawState = input as Partial<AppState> & {
    meta?: Partial<AppMeta>
    notebooks?: Array<
      Partial<Notebook> & {
        sectionGroups?: Array<
          Partial<SectionGroup> & {
            sections?: Array<Partial<Section> & { pages?: Array<Partial<Page>> }>
          }
        >
        sections?: Array<Partial<Section> & { pages?: Array<Partial<Page>> }>
      }
    >
  }

  const notebooks = (rawState.notebooks ?? [])
    .map((notebook, notebookIndex): Notebook | null => {
      if (!notebook?.id || !notebook?.name) {
        return null
      }

      const normalizePages = (
        rawPages: Array<Partial<Page>>,
        seed: number,
      ): Page[] =>
        rawPages
          .map((page, pageIndex): Page | null => {
            if (!page?.id || !page?.title) {
              return null
            }

            const content = typeof page.content === 'string' ? page.content : '<p></p>'
            const createdAt = typeof page.createdAt === 'string' ? page.createdAt : new Date().toISOString()
            const updatedAt = typeof page.updatedAt === 'string' ? page.updatedAt : createdAt

            return {
              accent:
                typeof page.accent === 'string'
                  ? page.accent
                  : accentPalette[(seed + pageIndex) % accentPalette.length],
              children: normalizePages(page.children ?? [], seed + pageIndex + 1),
              content,
              createdAt,
              id: page.id,
              inkStrokes: Array.isArray(page.inkStrokes)
                ? page.inkStrokes
                    .filter(
                      (stroke): stroke is InkStroke =>
                        !!stroke &&
                        typeof stroke === 'object' &&
                        typeof stroke.id === 'string' &&
                        typeof stroke.color === 'string' &&
                        typeof stroke.width === 'number' &&
                        Array.isArray(stroke.points),
                    )
                    .map((stroke) => ({
                      ...stroke,
                      points: stroke.points.filter(
                        (point): point is InkPoint =>
                          !!point &&
                          typeof point === 'object' &&
                          typeof point.x === 'number' &&
                          typeof point.y === 'number',
                      ),
                    }))
                : [],
              isCollapsed: typeof page.isCollapsed === 'boolean' ? page.isCollapsed : false,
              snippet:
                typeof page.snippet === 'string' && page.snippet.trim()
                  ? page.snippet
                  : buildSnippet(page.title, content),
              tags: Array.isArray(page.tags)
                ? page.tags.filter((tag): tag is string => typeof tag === 'string' && tag.trim().length > 0)
                : [],
              task:
                page.task &&
                typeof page.task === 'object' &&
                (page.task.status === 'open' || page.task.status === 'done')
                  ? {
                      dueAt: typeof page.task.dueAt === 'string' ? page.task.dueAt : null,
                      status: page.task.status,
                    }
                  : null,
              title: page.title,
              updatedAt,
            }
          })
          .filter((page): page is Page => page !== null)

      const normalizeSections = (
        rawSections: Array<Partial<Section> & { pages?: Array<Partial<Page>> }>,
        groupIndex: number,
      ) =>
        rawSections
          .map((section, sectionIndex): Section | null => {
            if (!section?.id || !section?.name) {
              return null
            }

            const pages = normalizePages(
              section.pages ?? [],
              notebookIndex + groupIndex + sectionIndex,
            )

            return {
              color:
                typeof section.color === 'string'
                  ? section.color
                  : accentPalette[(notebookIndex + groupIndex + sectionIndex) % accentPalette.length],
              id: section.id,
              name: section.name,
              pages,
              passwordHash: typeof section.passwordHash === 'string' ? section.passwordHash : null,
              passwordHint: typeof section.passwordHint === 'string' ? section.passwordHint : '',
            }
          })
          .filter((section): section is Section => section !== null)

      const rawGroups =
        notebook.sectionGroups && notebook.sectionGroups.length > 0
          ? notebook.sectionGroups
          : [
              {
                id: createId(),
                isCollapsed: false,
                name: 'Sections',
                sections:
                  (notebook as { sections?: Array<Partial<Section> & { pages?: Array<Partial<Page>> }> })
                    .sections ?? [],
              },
            ]

      const sectionGroups = rawGroups
        .map((group, groupIndex): SectionGroup | null => {
          if (!group?.id || !group?.name) {
            return null
          }

          return {
            id: group.id,
            isCollapsed: typeof group.isCollapsed === 'boolean' ? group.isCollapsed : false,
            name: group.name,
            sections: normalizeSections(group.sections ?? [], groupIndex),
          }
        })
        .filter((group): group is SectionGroup => group !== null)

      return {
        color:
          typeof notebook.color === 'string'
            ? notebook.color
            : accentPalette[notebookIndex % accentPalette.length],
        icon: typeof notebook.icon === 'string' ? notebook.icon : 'folder',
        id: notebook.id,
        name: notebook.name,
        sectionGroups,
      }
    })
    .filter((notebook): notebook is Notebook => notebook !== null)

  if (notebooks.length === 0) {
    return createStarterState()
  }

  return ensureSelection({
    meta: {
      assets:
        rawState.meta?.assets && typeof rawState.meta.assets === 'object'
          ? Object.fromEntries(
              Object.entries(rawState.meta.assets).filter(
                (entry): entry is [string, AppAsset] =>
                  typeof entry[0] === 'string' &&
                  !!entry[1] &&
                  typeof entry[1] === 'object' &&
                  typeof entry[1].id === 'string' &&
                  typeof entry[1].dataUrl === 'string' &&
                  typeof entry[1].name === 'string',
              ),
            )
          : {},
      pageSortMode:
        rawState.meta?.pageSortMode &&
        ['manual', 'updated-desc', 'updated-asc', 'title-asc', 'title-desc', 'created-desc'].includes(
          rawState.meta.pageSortMode,
        )
          ? rawState.meta.pageSortMode
          : defaultAppMeta().pageSortMode,
      pageVersions:
        rawState.meta?.pageVersions && typeof rawState.meta.pageVersions === 'object'
          ? Object.fromEntries(
              Object.entries(rawState.meta.pageVersions).map(([pageId, versions]) => [
                pageId,
                Array.isArray(versions)
                  ? versions
                      .filter(
                        (version): version is PageVersion =>
                          !!version &&
                          typeof version === 'object' &&
                          typeof version.id === 'string' &&
                          typeof version.title === 'string' &&
                          typeof version.content === 'string' &&
                          typeof version.savedAt === 'string',
                      )
                      .slice(0, 20)
                  : [],
              ]),
            )
          : {},
      recentPageIds: Array.isArray(rawState.meta?.recentPageIds)
        ? rawState.meta?.recentPageIds.filter((item): item is string => typeof item === 'string').slice(0, 8)
        : [],
      searchScope:
        rawState.meta?.searchScope && ['section', 'notebook', 'all'].includes(rawState.meta.searchScope)
          ? rawState.meta.searchScope
          : defaultAppMeta().searchScope,
    },
    notebooks,
    selectedNotebookId:
      typeof rawState.selectedNotebookId === 'string' ? rawState.selectedNotebookId : notebooks[0].id,
    selectedSectionGroupId:
      typeof rawState.selectedSectionGroupId === 'string' ? rawState.selectedSectionGroupId : '',
    selectedPageId: typeof rawState.selectedPageId === 'string' ? rawState.selectedPageId : '',
    selectedSectionId: typeof rawState.selectedSectionId === 'string' ? rawState.selectedSectionId : '',
  })
}

const getSelection = (state: AppState) => {
  const notebook =
    state.notebooks.find((item) => item.id === state.selectedNotebookId) ?? state.notebooks[0]
  const sectionGroup =
    notebook?.sectionGroups.find((item) => item.id === state.selectedSectionGroupId) ??
    notebook?.sectionGroups[0]
  const section =
    sectionGroup?.sections.find((item) => item.id === state.selectedSectionId) ?? sectionGroup?.sections[0]
  const page = section ? findPageById(section.pages, state.selectedPageId) ?? section.pages[0] : undefined

  return { notebook, page, section, sectionGroup }
}

const ensureSelection = (state: AppState): AppState => {
  const notebook = state.notebooks.find((item) => item.id === state.selectedNotebookId) ?? state.notebooks[0]
  const sectionGroup =
    notebook?.sectionGroups.find((item) => item.id === state.selectedSectionGroupId) ??
    notebook?.sectionGroups[0]
  const section =
    sectionGroup?.sections.find((item) => item.id === state.selectedSectionId) ?? sectionGroup?.sections[0]
  const page = section ? findPageById(section.pages, state.selectedPageId) ?? section.pages[0] : undefined

  if (!notebook || !sectionGroup || !section || !page) {
    return createStarterState()
  }

  return {
    ...state,
    selectedNotebookId: notebook.id,
    selectedSectionGroupId: sectionGroup.id,
    selectedPageId: page.id,
    selectedSectionId: section.id,
  }
}

function App() {
  const [appState, setAppState] = useState<AppState>(createStarterState)
  const [appInfo, setAppInfo] = useState<DesktopAppInfo | null>(null)
  const [activeTab, setActiveTab] = useState<RibbonTab>('Home')
  const [drawColor, setDrawColor] = useState('#1a73d9')
  const [isLoaded, setIsLoaded] = useState(false)
  const [query, setQuery] = useState('')
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null)
  const [isDirty, setIsDirty] = useState(false)
  const [isRecordingAudio, setIsRecordingAudio] = useState(false)
  const [isTranscribing, setIsTranscribing] = useState(false)
  const [unlockedSectionIds, setUnlockedSectionIds] = useState<string[]>([])
  const [saveLabel, setSaveLabel] = useState('Loading notes...')
  const [isCheckingForUpdates, setIsCheckingForUpdates] = useState(false)
  const [isStyleMenuOpen, setIsStyleMenuOpen] = useState(false)
  const [isFontMenuOpen, setIsFontMenuOpen] = useState(false)
  const [isFontSizeMenuOpen, setIsFontSizeMenuOpen] = useState(false)
  const [isCopilotOpen, setIsCopilotOpen] = useState(false)
  const [selectedFontFamily, setSelectedFontFamily] = useState('Calibri')
  const [selectedFontSize, setSelectedFontSize] = useState('11')
  const [dragState, setDragState] = useState<DragState | null>(null)
  const [dropTarget, setDropTarget] = useState<DropTarget | null>(null)
  const [dragPosition, setDragPosition] = useState<DragPosition | null>(null)
  const [recentNotebookEntries, setRecentNotebookEntries] = useState<RecentNotebookEntry[]>([])
  const editorRef = useRef<HTMLDivElement | null>(null)
  const imageInputRef = useRef<HTMLInputElement | null>(null)
  const attachmentInputRef = useRef<HTMLInputElement | null>(null)
  const printoutInputRef = useRef<HTMLInputElement | null>(null)
  const isCheckingForUpdatesRef = useRef(false)
  const titlebarSearchRef = useRef<HTMLDivElement | null>(null)
  const searchInputRef = useRef<HTMLInputElement | null>(null)
  const drawSurfaceRef = useRef<SVGSVGElement | null>(null)
  const mediaRecorderRef = useRef<MediaRecorder | null>(null)
  const speechRecognitionRef = useRef<BrowserSpeechRecognition | null>(null)
  const speechTranscriptRef = useRef('')
  const recordingChunksRef = useRef<Blob[]>([])
  const inkDrawingRef = useRef<InkStroke | null>(null)
  const styleMenuRef = useRef<HTMLDivElement | null>(null)
  const fontMenuRef = useRef<HTMLDivElement | null>(null)
  const fontSizeMenuRef = useRef<HTMLDivElement | null>(null)
  const searchResultsPanelRef = useRef<HTMLDivElement | null>(null)
  const styleMenuPanelRef = useRef<HTMLDivElement | null>(null)
  const fontMenuPanelRef = useRef<HTMLDivElement | null>(null)
  const fontSizeMenuPanelRef = useRef<HTMLDivElement | null>(null)
  const saveTimer = useRef<number | null>(null)
  const suppressClickAfterDragRef = useRef(false)
  const selectionRangeRef = useRef<Range | null>(null)
  const lastSavedPayloadRef = useRef('')
  const trackedRecentPageRef = useRef('')
  const shortcutActionsRef = useRef({
    addPage: () => {},
    createNotebook: () => {},
    demoteCurrentPage: () => {},
    openNotebook: () => {},
    promoteCurrentPage: () => {},
    saveNow: async () => {},
    saveNotebookAs: async () => {},
  })

  const { notebook, page, section, sectionGroup } = useMemo(() => getSelection(appState), [appState])
  const pageSortMode = appState.meta.pageSortMode
  const searchScope = appState.meta.searchScope
  const isCurrentSectionLocked = Boolean(section?.passwordHash && !unlockedSectionIds.includes(section.id))

  useEffect(() => {
    const load = async () => {
      try {
        setRecentNotebookEntries(loadRecentNotebookEntries())
        const [rawData, info] = await Promise.all([loadDesktopData(), getDesktopAppInfo()])
        if (info) setAppInfo(info)
        let nextState = rawData ? normalizeAppState(JSON.parse(rawData)) : createStarterState()
        const lastOpenedPath = loadLastOpenedNotebookPath()
        if (lastOpenedPath) {
          try {
            const rawNotebook = await openNotebookDirectory(lastOpenedPath)
            nextState = mergeNotebookIntoState(nextState, JSON.parse(rawNotebook) as Notebook)
          } catch {
            localStorage.removeItem(LAST_OPENED_NOTEBOOK_KEY)
          }
        }
        setAppState(nextState)
        lastSavedPayloadRef.current = JSON.stringify(nextState)
        trackedRecentPageRef.current = nextState.selectedPageId
        setIsDirty(false)
        setSaveLabel('All changes saved')
      } catch {
        setSaveLabel('Loaded sample notebook')
      } finally {
        setIsLoaded(true)
      }
    }
    void load()
  }, [])

  const runUpdateCheck = async (mode: 'automatic' | 'manual' = 'manual') => {
    if (isCheckingForUpdatesRef.current) return

    isCheckingForUpdatesRef.current = true
    setIsCheckingForUpdates(true)

    try {
      const update = await checkForDesktopUpdate()
      if (!update) {
        if (mode === 'manual') setSaveLabel('OnePlace is up to date')
        return
      }

      const shouldInstall = window.confirm(
        [
          `OnePlace ${update.version} is available.`,
          '',
          update.body?.trim() || 'An application update is ready to install.',
          '',
          'Install now? The app will restart after the update.',
        ].join('\n'),
      )

      if (!shouldInstall) {
        setSaveLabel(`Update available: ${update.version}`)
        return
      }

      let downloaded = 0
      let contentLength = 0
      setSaveLabel(`Downloading update ${update.version}...`)
      await downloadAndInstallDesktopUpdate((event) => {
        if (event.event === 'Started') {
          contentLength = event.data.contentLength ?? 0
          downloaded = 0
          setSaveLabel(`Downloading update ${update.version}...`)
          return
        }

        if (event.event === 'Progress') {
          downloaded += event.data.chunkLength
          if (contentLength > 0) {
            const percent = Math.min(100, Math.round((downloaded / contentLength) * 100))
            setSaveLabel(`Installing update ${update.version}... ${percent}%`)
          }
          return
        }

        setSaveLabel(`Restarting into OnePlace ${update.version}...`)
      })
    } catch {
      setSaveLabel('Update check failed')
    } finally {
      isCheckingForUpdatesRef.current = false
      setIsCheckingForUpdates(false)
    }
  }

  useEffect(() => {
    if (!isLoaded) return
    let cancelled = false

    const checkForUpdatesOnLaunch = async () => {
      if (isCheckingForUpdatesRef.current) return

      isCheckingForUpdatesRef.current = true
      setIsCheckingForUpdates(true)

      try {
        const update = await checkForDesktopUpdate()
        if (!update || cancelled) return

        const shouldInstall = window.confirm(
          [
            `OnePlace ${update.version} is available.`,
            '',
            update.body?.trim() || 'An application update is ready to install.',
            '',
            'Install now? The app will restart after the update.',
          ].join('\n'),
        )

        if (!shouldInstall) {
          setSaveLabel(`Update available: ${update.version}`)
          return
        }

        let downloaded = 0
        let contentLength = 0
        setSaveLabel(`Downloading update ${update.version}...`)
        await downloadAndInstallDesktopUpdate((event) => {
          if (cancelled) return

          if (event.event === 'Started') {
            contentLength = event.data.contentLength ?? 0
            downloaded = 0
            setSaveLabel(`Downloading update ${update.version}...`)
            return
          }

          if (event.event === 'Progress') {
            downloaded += event.data.chunkLength
            if (contentLength > 0) {
              const percent = Math.min(100, Math.round((downloaded / contentLength) * 100))
              setSaveLabel(`Installing update ${update.version}... ${percent}%`)
            }
            return
          }

          setSaveLabel(`Restarting into OnePlace ${update.version}...`)
        })
      } catch {
        if (!cancelled) setSaveLabel('Update check failed')
      } finally {
        isCheckingForUpdatesRef.current = false
        if (!cancelled) setIsCheckingForUpdates(false)
      }
    }

    void checkForUpdatesOnLaunch()

    return () => {
      cancelled = true
    }
  }, [isLoaded])

  useEffect(() => {
    if (!isLoaded) return
    const payload = JSON.stringify(appState)
    if (payload === lastSavedPayloadRef.current) {
      setIsDirty(false)
      return
    }
    if (saveTimer.current) window.clearTimeout(saveTimer.current)
    setIsDirty(true)
    setSaveLabel('Saving...')
    saveTimer.current = window.setTimeout(() => {
      void saveDesktopData(payload).then((result) => {
        lastSavedPayloadRef.current = payload
        setIsDirty(false)
        setSaveLabel(`Saved ${formatDate(result.savedAt)}`)
      })
    }, 250)
    return () => {
      if (saveTimer.current) window.clearTimeout(saveTimer.current)
    }
  }, [appState, isLoaded])

  useEffect(() => {
    if (editorRef.current && page) {
      const hydrated = hydratePageContent(page.content, appState.meta.assets)
      if (editorRef.current.innerHTML !== hydrated) {
        editorRef.current.innerHTML = hydrated
      }
    }
  }, [appState.meta.assets, page])

  useEffect(() => {
    if (!isLoaded || !appState.selectedPageId || trackedRecentPageRef.current === appState.selectedPageId) return

    trackedRecentPageRef.current = appState.selectedPageId
    setAppState((current) => {
      if (current.meta.recentPageIds[0] === current.selectedPageId) return current

      return {
        ...current,
        meta: {
          ...current.meta,
          recentPageIds: [current.selectedPageId, ...current.meta.recentPageIds.filter((id) => id !== current.selectedPageId)].slice(0, 8),
        },
      }
    })
  }, [appState.selectedPageId, isLoaded])

  useEffect(() => {
    const pageTitle = page?.title?.trim() || 'Untitled Page'
    const notebookTitle = notebook?.name?.trim() || 'Notebook'
    const nextTitle = `${isDirty ? '* ' : ''}${pageTitle} - ${notebookTitle} - OneNote`
    document.title = nextTitle
  }, [isDirty, notebook?.name, page?.title])

  useEffect(() => {
    const warnBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!isDirty) return
      event.preventDefault()
      event.returnValue = ''
    }

    window.addEventListener('beforeunload', warnBeforeUnload)
    return () => {
      window.removeEventListener('beforeunload', warnBeforeUnload)
    }
  }, [isDirty])

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

  useEffect(() => {
    const captureSelection = () => {
      const selection = window.getSelection()
      if (!selection || selection.rangeCount === 0) return
      const range = selection.getRangeAt(0)
      if (editorRef.current?.contains(range.commonAncestorContainer)) {
        selectionRangeRef.current = range.cloneRange()
      }
    }

    document.addEventListener('selectionchange', captureSelection)
    return () => {
      document.removeEventListener('selectionchange', captureSelection)
    }
  }, [])

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
  }, [])

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

  const matchesSearch = (
    targetPage: Page,
    notebookName: string,
    groupName: string,
    sectionName: string,
    needle: string,
  ) =>
    `${targetPage.title} ${targetPage.snippet} ${targetPage.content.replace(/<[^>]+>/g, ' ')} ${targetPage.tags.join(' ')} ${targetPage.task?.status ?? ''} ${targetPage.task?.dueAt ?? ''} ${notebookName} ${groupName} ${sectionName}`
      .toLowerCase()
      .includes(needle)

  const visiblePages = useMemo(() => {
    if (!section || isCurrentSectionLocked) return []
    const sortedPages = sortPagesTree(section.pages, pageSortMode)
    const needle = query.trim().toLowerCase()
    const flattened = flattenPages(sortedPages, 0, Boolean(needle && searchScope === 'section'))
    if (!needle || searchScope !== 'section') return flattened

    return flattened.filter(({ page: item }) =>
      matchesSearch(item, notebook?.name ?? '', sectionGroup?.name ?? '', section.name, needle),
    )
  }, [isCurrentSectionLocked, notebook?.name, pageSortMode, query, searchScope, section, sectionGroup?.name])

  const searchResults = useMemo(() => {
    const needle = query.trim().toLowerCase()
    if (!needle) return []

    const notebookTargets =
      searchScope === 'section'
        ? (notebook && sectionGroup && section ? [notebook] : [])
        : searchScope === 'notebook' && notebook
          ? [notebook]
          : appState.notebooks

    const results = notebookTargets.flatMap((entry) =>
      entry.sectionGroups.flatMap((group) =>
        group.sections.flatMap((part) => {
          if (
            searchScope === 'section' &&
            (entry.id !== notebook?.id || group.id !== sectionGroup?.id || part.id !== section?.id)
          ) {
            return []
          }

          return flattenPages(sortPagesTree(part.pages, pageSortMode), 0, true)
            .filter((item) => matchesSearch(item.page, entry.name, group.name, part.name, needle))
            .map(
              (item): SearchResult => ({
                groupId: group.id,
                groupName: group.name,
                isSubpage: item.depth > 0,
                notebookId: entry.id,
                notebookName: entry.name,
                page: item.page,
                sectionId: part.id,
                sectionName: part.name,
              }),
            )
        }),
      ),
    )

    return results.sort(
      (left, right) => new Date(right.page.updatedAt).getTime() - new Date(left.page.updatedAt).getTime(),
    )
  }, [appState.notebooks, notebook, pageSortMode, query, searchScope, section, sectionGroup])

  const recentPages = useMemo(
    () =>
      appState.meta.recentPageIds
        .map((pageId) =>
          appState.notebooks.flatMap((entry) =>
            entry.sectionGroups.flatMap((group) =>
              group.sections.flatMap((part) =>
                flattenPages(part.pages, 0, true)
                  .filter((item) => item.page.id === pageId)
                  .map((item) => ({
                    groupId: group.id,
                    groupName: group.name,
                    notebookId: entry.id,
                    notebookName: entry.name,
                    page: item.page,
                    sectionId: part.id,
                    sectionName: part.name,
                  })),
              ),
            ),
          )[0] ?? null,
        )
        .filter(
          (
            item,
          ): item is Omit<SearchResult, 'isSubpage'> & {
            page: Page
          } => item !== null,
        ),
    [appState.meta.recentPageIds, appState.notebooks],
  )

  const selectedPageLocation = useMemo(
    () => (section && page ? findPageLocation(section.pages, page.id) : undefined),
    [page, section],
  )
  const canDeleteNotebook = appState.notebooks.length > 1
  const canDeleteSectionGroup = Boolean(notebook && notebook.sectionGroups.length > 1)
  const canDeleteSection = Boolean(sectionGroup && sectionGroup.sections.length > 1)
  const canDeletePage = Boolean(section && countPages(section.pages) > 1)
  const canPromotePage = Boolean(selectedPageLocation?.parentId)
  const canDemotePage = Boolean(selectedPageLocation && selectedPageLocation.index > 0)
  const notebookSectionCount =
    notebook?.sectionGroups.reduce((total, group) => total + group.sections.length, 0) ?? 0
  const notebookPageCount =
    notebook?.sectionGroups.reduce(
      (total, group) => total + group.sections.reduce((sectionTotal, entry) => sectionTotal + countPages(entry.pages), 0),
      0,
    ) ?? 0
  const selectNotebook = (notebookId: string) => {
    setAppState((current) => {
      const nextNotebook = current.notebooks.find((item) => item.id === notebookId)
      const nextGroup = nextNotebook?.sectionGroups[0]
      const nextSection = nextGroup?.sections[0]
      const nextPage = nextSection?.pages[0]
      if (!nextNotebook) return current
      if (!nextGroup || !nextSection || !nextPage) {
        return { ...current, selectedNotebookId: nextNotebook.id }
      }
      return {
        ...current,
        selectedNotebookId: nextNotebook.id,
        selectedSectionGroupId: nextGroup.id,
        selectedPageId: nextPage.id,
        selectedSectionId: nextSection.id,
      }
    })
  }

  const trackRecentNotebook = (path: string, name: string) => {
    const nextEntries = [{ name, path }, ...recentNotebookEntries.filter((item) => item.path !== path)].slice(0, 8)
    setRecentNotebookEntries(nextEntries)
    saveRecentNotebookEntries(nextEntries)
    saveLastOpenedNotebookPath(path)
  }

  const loadNotebookFromPath = async (path: string) => {
    const rawNotebook = await openNotebookDirectory(path)
    const openedNotebook = JSON.parse(rawNotebook) as Notebook

    setAppState((current) => mergeNotebookIntoState(current, openedNotebook))
    trackRecentNotebook(path, openedNotebook.name)
    setActiveTab('Home')
    setSaveLabel(`Opened ${openedNotebook.name}`)
  }

  const openNotebook = () => {
    void (async () => {
      try {
        const path = await pickNotebookDirectory()
        if (!path) return
        await loadNotebookFromPath(path)
      } catch {
        window.alert('That folder does not contain a valid notebook.')
      }
    })()
  }

  const selectSection = (groupId: string, sectionId: string) => {
    setAppState((current) => {
      const currentNotebook = current.notebooks.find((item) => item.id === current.selectedNotebookId)
      const nextGroup = currentNotebook?.sectionGroups.find((item) => item.id === groupId)
      const nextSection = nextGroup?.sections.find((item) => item.id === sectionId)
      const nextPage = nextSection?.pages[0]
      if (!nextSection || !nextPage) return current
      return {
        ...current,
        selectedPageId: nextPage.id,
        selectedSectionGroupId: nextGroup?.id ?? current.selectedSectionGroupId,
        selectedSectionId: nextSection.id,
      }
    })
  }

  const selectPage = (pageId: string) => {
    setAppState((current) => ({ ...current, selectedPageId: pageId }))
  }

  const openSearchResult = (result: SearchResult) => {
    setAppState((current) => ({
      ...current,
      selectedNotebookId: result.notebookId,
      selectedSectionGroupId: result.groupId,
      selectedSectionId: result.sectionId,
      selectedPageId: result.page.id,
    }))
    setQuery('')
  }

  const setPageSortMode = (nextMode: PageSortMode) => {
    setAppState((current) => ({
      ...current,
      meta:
        current.meta.pageSortMode === nextMode
          ? current.meta
          : {
              ...current.meta,
              pageSortMode: nextMode,
            },
    }))
  }

  const openLinkedPage = (pageId: string) => {
    const match = appState.notebooks.flatMap((entry) =>
      entry.sectionGroups.flatMap((group) =>
        group.sections.flatMap((part) =>
          flattenPages(part.pages, 0, true)
            .filter((candidate) => candidate.page.id === pageId)
            .map((candidate) => ({
              groupId: group.id,
              notebookId: entry.id,
              pageId: candidate.page.id,
              sectionId: part.id,
            })),
        ),
      ),
    )[0]

    if (!match) return

    setAppState((current) => ({
      ...current,
      selectedNotebookId: match.notebookId,
      selectedSectionGroupId: match.groupId,
      selectedSectionId: match.sectionId,
      selectedPageId: match.pageId,
    }))
  }

  const lockSection = (sectionId: string) => {
    setUnlockedSectionIds((current) => current.filter((item) => item !== sectionId))
  }

  const unlockSection = async (sectionId: string) => {
    const targetSection = appState.notebooks
      .flatMap((entry) => entry.sectionGroups.flatMap((group) => group.sections))
      .find((entry) => entry.id === sectionId)
    if (!targetSection?.passwordHash) return

    const promptLabel = targetSection.passwordHint
      ? `Enter password for ${targetSection.name}\nHint: ${targetSection.passwordHint}`
      : `Enter password for ${targetSection.name}`
    const password = window.prompt(promptLabel, '') ?? ''
    if (!password) return
    const hash = await hashSecret(password)
    if (hash !== targetSection.passwordHash) {
      window.alert('Incorrect password.')
      return
    }

    setUnlockedSectionIds((current) => [...new Set([...current, sectionId])])
  }

  const protectSection = async (groupId: string, sectionId: string) => {
    const password = window.prompt('Set a password for this section', '')?.trim()
    if (!password) return
    const hint = window.prompt('Password hint (optional)', '')?.trim() ?? ''
    const hash = await hashSecret(password)
    setAppState((current) => ({
      ...current,
      notebooks: current.notebooks.map((item) => ({
        ...item,
        sectionGroups: item.sectionGroups.map((group) => ({
          ...group,
          sections: group.sections.map((entry) =>
            entry.id === sectionId && group.id === groupId
              ? {
                  ...entry,
                  passwordHash: hash,
                  passwordHint: hint,
                }
              : entry,
          ),
        })),
      })),
    }))
    lockSection(sectionId)
    setSaveLabel('Section protection enabled')
  }

  const removeSectionProtection = async (groupId: string, sectionId: string) => {
    const targetSection = appState.notebooks
      .flatMap((entry) => entry.sectionGroups.flatMap((group) => group.sections))
      .find((entry) => entry.id === sectionId)
    if (!targetSection?.passwordHash) return

    const password = window.prompt('Enter the current password to remove protection', '') ?? ''
    if (!password) return
    const hash = await hashSecret(password)
    if (hash !== targetSection.passwordHash) {
      window.alert('Incorrect password.')
      return
    }

    setAppState((current) => ({
      ...current,
      notebooks: current.notebooks.map((item) => ({
        ...item,
        sectionGroups: item.sectionGroups.map((group) => ({
          ...group,
          sections: group.sections.map((entry) =>
            entry.id === sectionId && group.id === groupId
              ? {
                  ...entry,
                  passwordHash: null,
                  passwordHint: '',
                }
              : entry,
          ),
        })),
      })),
    }))
    setUnlockedSectionIds((current) => current.filter((item) => item !== sectionId))
    setSaveLabel('Section protection removed')
  }

  const addTagToCurrentPage = () => {
    if (!page) return
    const tag = window.prompt('Add a tag', '')?.trim()
    if (!tag) return
    const nextTags = [...new Set([...page.tags, tag])]
    updatePage({ tags: nextTags })
  }

  const toggleCurrentTask = () => {
    if (!page) return
    updatePage({ task: page.task ? null : defaultTask() })
  }

  const toggleCurrentTaskComplete = () => {
    if (!page?.task) return
    updatePage({
      task: {
        ...page.task,
        status: page.task.status === 'done' ? 'open' : 'done',
      },
    })
  }

  const setCurrentTaskDueDate = () => {
    if (!page) return
    const currentDue = page.task?.dueAt?.slice(0, 10) ?? ''
    const value = window.prompt('Due date (YYYY-MM-DD), leave blank to clear', currentDue)?.trim()
    if (value === undefined) return
    updatePage({
      task: page.task
        ? {
            ...page.task,
            dueAt: value ? new Date(value).toISOString() : null,
          }
        : {
            ...defaultTask(),
            dueAt: value ? new Date(value).toISOString() : null,
          },
    })
  }

  const renamePage = (pageId: string) => {
    if (!section) return
    const currentPage = findPageById(section.pages, pageId)
    if (!currentPage) return
    const nextName = window.prompt('Rename page', currentPage.title)?.trim()
    if (!nextName || nextName === currentPage.title) return

    setAppState((current) => ({
      ...current,
      notebooks: current.notebooks.map((item) => ({
        ...item,
        sectionGroups: item.sectionGroups.map((group) => ({
          ...group,
          sections: group.sections.map((part) =>
            part.id === current.selectedSectionId
              ? {
                  ...part,
                  pages: updateNestedPages(part.pages, pageId, (note) => ({
                    ...note,
                    snippet: buildSnippet(nextName, note.content),
                    title: nextName,
                    updatedAt: new Date().toISOString(),
                  })),
                }
              : part,
          ),
        })),
      })),
    }))
  }

  const saveCurrentPageVersion = (targetPage = page) => {
    if (!targetPage) return
    setAppState((current) => ({
      ...current,
      meta: {
        ...current.meta,
        pageVersions: recordPageVersion(
          current.meta.pageVersions,
          targetPage.id,
          targetPage.title,
          targetPage.content,
        ),
      },
    }))
    setSaveLabel(`Saved version for ${targetPage.title}`)
  }

  const restoreSavedPageVersion = () => {
    if (!page) return
    const versions = appState.meta.pageVersions[page.id] ?? []
    if (versions.length === 0) {
      window.alert('No saved versions for this page yet.')
      return
    }

    const options = versions
      .map((version, index) => `${index + 1}. ${formatDate(version.savedAt)}${index === 0 ? ' (Latest)' : ''}`)
      .join('\n')
    const picked = window.prompt(`Restore which version?\n\n${options}`, '1')?.trim()
    if (!picked) return
    const version = versions[Number(picked) - 1]
    if (!version) return

    updatePage({
      content: version.content,
      title: version.title,
    })
    setSaveLabel(`Restored ${formatDate(version.savedAt)}`)
  }

  const openContextMenu = (
    event: ReactMouseEvent<HTMLElement>,
    items: ContextMenuItem[],
  ) => {
    event.preventDefault()
    setContextMenu({
      items,
      x: event.clientX,
      y: event.clientY,
    })
  }

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

  const setNotebookDropTarget = (event: ReactPointerEvent<HTMLElement>, notebookId: string) => {
    if (!dragState || dragState.type !== 'notebook') return
    event.preventDefault()
    event.stopPropagation()
    setDropTarget({ type: 'notebook', notebookId, position: getDropPosition(event) })
  }

  const moveNotebook = (targetNotebookId: string, position: 'before' | 'after') => {
    if (!dragState || dragState.type !== 'notebook' || dragState.notebookId === targetNotebookId) {
      return
    }

    setAppState((current) => ({
      ...current,
      notebooks: reorderItems(current.notebooks, dragState.notebookId, targetNotebookId, position),
    }))
    clearDragState(true)
  }

  const setSectionGroupDropTarget = (event: ReactPointerEvent<HTMLElement>, groupId: string) => {
    if (!dragState || dragState.type !== 'section-group') return
    event.preventDefault()
    event.stopPropagation()
    setDropTarget({ type: 'section-group', groupId, position: getDropPosition(event) })
  }

  const setSectionDropTarget = (
    event: ReactPointerEvent<HTMLElement>,
    groupId: string,
    sectionId: string,
  ) => {
    if (!dragState || dragState.type !== 'section') return
    event.preventDefault()
    event.stopPropagation()
    setDropTarget({ type: 'section', groupId, sectionId, position: getDropPosition(event) })
  }

  const setSectionGroupInsideDropTarget = (event: ReactPointerEvent<HTMLElement>, groupId: string) => {
    if (!dragState || dragState.type !== 'section') return
    event.preventDefault()
    setDropTarget({ type: 'section', groupId, position: 'inside' })
  }

  const setPageDropTarget = (event: ReactPointerEvent<HTMLElement>, pageId: string) => {
    if (!dragState || dragState.type !== 'page') return
    const draggedPage = section ? findPageById(section.pages, dragState.pageId) : undefined
    if (!draggedPage || dragState.pageId === pageId || pageContainsId(draggedPage, pageId)) {
      setDropTarget(null)
      return
    }

    event.preventDefault()
    event.stopPropagation()
    setDropTarget({ type: 'page', pageId, position: getDropPosition(event) })
  }

  const moveSectionGroup = (targetGroupId: string, position: 'before' | 'after') => {
    if (!notebook || !dragState || dragState.type !== 'section-group' || dragState.groupId === targetGroupId) {
      return
    }

    setAppState((current) => ({
      ...current,
      notebooks: current.notebooks.map((item) =>
        item.id === current.selectedNotebookId
          ? {
              ...item,
              sectionGroups: reorderItems(item.sectionGroups, dragState.groupId, targetGroupId, position),
            }
          : item,
      ),
    }))
    clearDragState(true)
  }

  const moveSection = (
    targetGroupId: string,
    targetSectionId: string,
    position: 'before' | 'after',
  ) => {
    if (!notebook || !dragState || dragState.type !== 'section') return
    if (dragState.groupId === targetGroupId && dragState.sectionId === targetSectionId) return

    setAppState((current) => {
      const currentNotebook = current.notebooks.find((item) => item.id === current.selectedNotebookId)
      if (!currentNotebook) return current

      const sourceGroup = currentNotebook.sectionGroups.find((group) => group.id === dragState.groupId)
      const draggedSection = sourceGroup?.sections.find((item) => item.id === dragState.sectionId)
      if (!sourceGroup || !draggedSection) return current

      return {
        ...current,
        notebooks: current.notebooks.map((item) => {
          if (item.id !== current.selectedNotebookId) return item

          let sectionToInsert: Section | null = null
          const strippedGroups = item.sectionGroups.map((group) => {
            if (group.id !== dragState.groupId) return group
            const remainingSections = group.sections.filter((entry) => {
              const keep = entry.id !== dragState.sectionId
              if (!keep) sectionToInsert = entry
              return keep
            })
            return { ...group, sections: remainingSections }
          })

          if (!sectionToInsert) return item

          return {
            ...item,
            sectionGroups: strippedGroups.map((group) => {
              if (group.id !== targetGroupId) return group
              return {
                ...group,
                sections: reorderItems(
                  [...group.sections, sectionToInsert as Section],
                  dragState.sectionId,
                  targetSectionId,
                  position,
                ),
              }
            }),
          }
        }),
        selectedSectionGroupId: targetGroupId,
        selectedSectionId: draggedSection.id,
      }
    })
    clearDragState(true)
  }

  const moveSectionToGroup = (targetGroupId: string) => {
    if (!notebook || !dragState || dragState.type !== 'section') return
    if (dragState.groupId === targetGroupId) return

    setAppState((current) => {
      const currentNotebook = current.notebooks.find((item) => item.id === current.selectedNotebookId)
      if (!currentNotebook) return current

      const sourceGroup = currentNotebook.sectionGroups.find((group) => group.id === dragState.groupId)
      const draggedSection = sourceGroup?.sections.find((item) => item.id === dragState.sectionId)
      if (!draggedSection) return current

      return {
        ...current,
        notebooks: current.notebooks.map((item) => {
          if (item.id !== current.selectedNotebookId) return item

          let sectionToMove: Section | null = null
          const groupsWithoutDraggedSection = item.sectionGroups.map((group) => {
            if (group.id !== dragState.groupId) return group
            return {
              ...group,
              sections: group.sections.filter((entry) => {
                const keep = entry.id !== dragState.sectionId
                if (!keep) sectionToMove = entry
                return keep
              }),
            }
          })

          if (!sectionToMove) return item

          return {
            ...item,
            sectionGroups: groupsWithoutDraggedSection.map((group) =>
              group.id === targetGroupId
                ? { ...group, sections: [...group.sections, sectionToMove as Section] }
                : group,
            ),
          }
        }),
        selectedSectionGroupId: targetGroupId,
        selectedSectionId: draggedSection.id,
      }
    })
    clearDragState(true)
  }

  const toggleSectionGroupCollapse = (groupId: string) => {
    setAppState((current) => ({
      ...current,
      notebooks: current.notebooks.map((item) =>
        item.id === current.selectedNotebookId
          ? {
              ...item,
              sectionGroups: item.sectionGroups.map((group) =>
                group.id === groupId ? { ...group, isCollapsed: !group.isCollapsed } : group,
              ),
            }
          : item,
      ),
    }))
  }

  const movePage = (targetPageId: string, position: 'before' | 'after') => {
    if (!section || !dragState || dragState.type !== 'page' || dragState.pageId === targetPageId) return

    setAppState((current) => {
      let movedPage: Page | undefined
      const nextState = {
        ...current,
        notebooks: current.notebooks.map((item) => ({
          ...item,
          sectionGroups: item.sectionGroups.map((group) => ({
            ...group,
            sections: group.sections.map((entry) => {
              if (entry.id !== current.selectedSectionId) return entry

              const removed = removePageById(entry.pages, dragState.pageId)
              movedPage = removed.page
              if (!movedPage || pageContainsId(movedPage, targetPageId)) {
                movedPage = undefined
                return entry
              }
              return {
                ...entry,
                pages: insertPageRelative(removed.pages, targetPageId, movedPage, position),
              }
            }),
          })),
        })),
        selectedPageId: dragState.pageId,
      }
      return movedPage ? nextState : current
    })
    clearDragState(true)
  }

  const togglePageCollapse = (pageId: string) => {
    setAppState((current) => {
      let nextSelectedPageId = current.selectedPageId

      return {
        ...current,
        notebooks: current.notebooks.map((item) => ({
          ...item,
          sectionGroups: item.sectionGroups.map((group) => ({
            ...group,
            sections: group.sections.map((entry) => {
              if (entry.id !== current.selectedSectionId) return entry

              return {
                ...entry,
                pages: updateNestedPages(entry.pages, pageId, (note) => {
                  const nextCollapsed = !note.isCollapsed
                  if (nextCollapsed && hasChildPageSelected(note, current.selectedPageId)) {
                    nextSelectedPageId = note.id
                  }

                  return {
                    ...note,
                    isCollapsed: nextCollapsed,
                  }
                }),
              }
            }),
          })),
        })),
        selectedPageId: nextSelectedPageId,
      }
    })
  }

  const updatePage = (updates: Partial<Page>) => {
    if (!page) return
    setAppState((current) => ({
      ...current,
      notebooks: current.notebooks.map((item) => ({
        ...item,
        sectionGroups: item.sectionGroups.map((group) => ({
          ...group,
          sections: group.sections.map((entry) => ({
            ...entry,
            pages: updateNestedPages(entry.pages, current.selectedPageId, (note) => {
              const nextTitle = typeof updates.title === 'string' ? updates.title : note.title
              const nextContent = typeof updates.content === 'string' ? updates.content : note.content
              const nextUpdatedAt = new Date().toISOString()
              return {
                ...note,
                ...updates,
                snippet: buildSnippet(nextTitle, nextContent, nextUpdatedAt),
                updatedAt: nextUpdatedAt,
              }
            }),
          })),
        })),
      })),
    }))
  }

  const focusEditor = (restoreSelection = true) => {
    editorRef.current?.focus()
    if (!restoreSelection || !selectionRangeRef.current) return

    const selection = window.getSelection()
    if (!selection) return
    selection.removeAllRanges()
    selection.addRange(selectionRangeRef.current)
  }

  const runEditorCommand = (command: string, value?: string) => {
    focusEditor()
    document.execCommand(command, false, value)
    window.setTimeout(() => {
      updatePage({
        content: dehydratePageContent(sanitizePastedHtml(editorRef.current?.innerHTML ?? page?.content ?? '')),
      })
    }, 0)
  }

  const syncEditorContent = () => {
    updatePage({
      content: dehydratePageContent(sanitizePastedHtml(editorRef.current?.innerHTML ?? page?.content ?? '')),
    })
  }

  const getActiveEditorRange = () => {
    const selection = window.getSelection()
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0)
      if (editorRef.current?.contains(range.commonAncestorContainer)) {
        return range.cloneRange()
      }
    }

    if (selectionRangeRef.current) {
      return selectionRangeRef.current.cloneRange()
    }

    if (!editorRef.current) return null

    const fallbackRange = document.createRange()
    fallbackRange.selectNodeContents(editorRef.current)
    fallbackRange.collapse(false)
    return fallbackRange
  }

  const insertHtmlAtSelection = (html: string) => {
    focusEditor()

    const range = getActiveEditorRange()
    if (!range) return

    const fragment = range.createContextualFragment(html)
    const lastNode = fragment.lastChild

    range.deleteContents()
    range.insertNode(fragment)

    const selection = window.getSelection()
    if (selection) {
      const nextRange = document.createRange()
      if (lastNode) {
        nextRange.setStartAfter(lastNode)
      } else {
        nextRange.selectNodeContents(editorRef.current!)
        nextRange.collapse(false)
      }
      nextRange.collapse(true)
      selection.removeAllRanges()
      selection.addRange(nextRange)
      selectionRangeRef.current = nextRange.cloneRange()
    }

    window.setTimeout(syncEditorContent, 0)
  }

  const insertTextAsHtml = (value: string) => {
    insertHtmlAtSelection(plainTextToHtml(value))
  }

  const insertChecklist = () => {
    insertHtmlAtSelection('<ul class="checklist"><li><label><input type="checkbox" /> New task</label></li></ul>')
  }

  const insertTable = () => {
    insertHtmlAtSelection(
      '<table class="action-table"><tbody><tr><th>Topic</th><th>Notes</th><th>Owner</th></tr><tr><td></td><td></td><td></td></tr></tbody></table><p></p>',
    )
  }

  const getSelectionContainerElement = () => {
    const selection = window.getSelection()
    if (!selection || selection.rangeCount === 0) return null
    const node = selection.getRangeAt(0).startContainer
    return node instanceof Element ? node : node.parentElement
  }

  const getChecklistContext = () => {
    const container = getSelectionContainerElement()
    const item = container?.closest('li')
    const list = container?.closest('ul.checklist')
    if (!item || !list) return null
    return { item, list }
  }

  const focusChecklistItem = (item: HTMLLIElement) => {
    const label = item.querySelector('label')
    const targetNode = label?.lastChild
    if (!label || !targetNode) return
    const range = document.createRange()
    const offset = targetNode.textContent?.length ?? 0
    range.setStart(targetNode, offset)
    range.collapse(true)
    const selection = window.getSelection()
    selection?.removeAllRanges()
    selection?.addRange(range)
    selectionRangeRef.current = range.cloneRange()
  }

  const createChecklistItemNode = (text = '') => {
    const item = document.createElement('li')
    const label = document.createElement('label')
    const input = document.createElement('input')
    input.type = 'checkbox'
    label.append(input, document.createTextNode(` ${text}`))
    item.append(label)
    return item
  }

  const insertExternalLink = () => {
    const urlInput = window.prompt('Enter a link URL', 'https://')
    const url = urlInput?.trim()
    if (!url) return
    const href = /^https?:\/\//i.test(url) || /^mailto:/i.test(url) ? url : `https://${url}`
    const selectedText = window.getSelection()?.toString().trim()
    const label = window.prompt('Link text', selectedText || href)?.trim() || href
    insertHtmlAtSelection(
      `<a href="${escapeAttribute(href)}" target="_blank" rel="noreferrer">${escapeHtml(label)}</a>`,
    )
  }

  const insertInternalPageLink = () => {
    const flattenedPages = appState.notebooks.flatMap((entry) =>
      entry.sectionGroups.flatMap((group) =>
        group.sections.flatMap((part) =>
          flattenPages(part.pages, 0, true).map(({ page: linkedPage }) => ({
            label: `${entry.name} / ${group.name} / ${part.name} / ${linkedPage.title}`,
            pageId: linkedPage.id,
            title: linkedPage.title,
          })),
        ),
      ),
    )
    if (flattenedPages.length === 0) return

    const options = flattenedPages.map((item, index) => `${index + 1}. ${item.label}`).join('\n')
    const picked = window.prompt(`Link to which page?\n\n${options}`, '1')?.trim()
    if (!picked) return
    const target = flattenedPages[Number(picked) - 1]
    if (!target) return
    insertHtmlAtSelection(
      `<a class="internal-page-link" data-page-id="${escapeAttribute(target.pageId)}" href="#page:${escapeAttribute(target.pageId)}">${escapeHtml(target.title)}</a>`,
    )
  }

  const applyPageTemplate = () => {
    const options = pageTemplates.map((template, index) => `${index + 1}. ${template.label}`).join('\n')
    const picked = window.prompt(`Choose a template\n\n${options}`, '1')?.trim()
    if (!picked) return
    const template = pageTemplates[Number(picked) - 1]
    if (!template) return
    insertHtmlAtSelection(template.html)
  }

  const saveNow = async () => {
    const payload = JSON.stringify(appState)
    if (payload === lastSavedPayloadRef.current) {
      setSaveLabel('All changes saved')
      setIsDirty(false)
      return
    }
    setSaveLabel('Saving...')
    const result = await saveDesktopData(payload)
    lastSavedPayloadRef.current = payload
    setIsDirty(false)
    setSaveLabel(`Saved ${formatDate(result.savedAt)}`)
  }

  const saveNotebookAs = async () => {
    if (!notebook) return

    const path = await pickNotebookDirectory()
    if (!path) return

    const result = await exportNotebookDirectory(path, JSON.stringify(notebook))
    trackRecentNotebook(result.path, notebook.name)
    setSaveLabel(`Saved notebook to ${result.path}`)
  }

  const copySelection = async () => {
    const selection = window.getSelection()?.toString() ?? ''
    const activeElement = document.activeElement
    const inputSelection =
      activeElement instanceof HTMLInputElement || activeElement instanceof HTMLTextAreaElement
        ? activeElement.value.slice(activeElement.selectionStart ?? 0, activeElement.selectionEnd ?? 0)
        : ''
    const rawText = inputSelection || selection || editorRef.current?.innerText || page?.title || ''
    const normalized = normalizeTerminalText(rawText)
    if (!normalized.trim()) return

    await navigator.clipboard.writeText(normalized)
    setSaveLabel('Copied as terminal-safe text')
  }

  const readFileAsDataUrl = (file: File) =>
    new Promise<string>((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '')
      reader.onerror = () => reject(reader.error)
      reader.readAsDataURL(file)
    })

  const addAssetsToState = (assets: AppAsset[]) => {
    if (assets.length === 0) return

    setAppState((current) => ({
      ...current,
      meta: {
        ...current.meta,
        assets: {
          ...current.meta.assets,
          ...Object.fromEntries(assets.map((asset) => [asset.id, asset])),
        },
      },
    }))
  }

  const openImagePicker = () => {
    imageInputRef.current?.click()
  }

  const openAttachmentPicker = () => {
    attachmentInputRef.current?.click()
  }

  const openPrintoutPicker = () => {
    printoutInputRef.current?.click()
  }

  const emailCurrentPage = () => {
    if (!page) return

    const plainContent = extractSnippetText(hydratePageContent(page.content, appState.meta.assets))
    const body = [
      `Notebook: ${notebook?.name ?? 'Notebook'}`,
      `Section: ${section?.name ?? 'Section'}`,
      '',
      plainContent || page.title,
    ]
      .join('\n')
      .slice(0, 1800)

    window.location.href = `mailto:?subject=${encodeURIComponent(page.title)}&body=${encodeURIComponent(body)}`
    setSaveLabel('Opened default mail client')
  }

  const insertTranscriptIntoPage = (transcript: string) => {
    if (!page) return

    const block = `
      <section class="template-block">
        <h3>Transcript</h3>
        ${plainTextToHtml(transcript)}
      </section>
    `

    updatePage({
      content: `${page.content}${block}`,
    })
  }

  const insertAudioNote = async (blob: Blob) => {
    const assetId = createId()
    const extension = blob.type.split('/')[1] || 'webm'
    const file = new File([blob], `audio-note-${new Date().toISOString().slice(0, 10)}.${extension}`, {
      type: blob.type || 'audio/webm',
    })
    const dataUrl = await readFileAsDataUrl(file)
    addAssetsToState([
      {
        createdAt: new Date().toISOString(),
        dataUrl,
        id: assetId,
        kind: 'audio',
        mimeType: file.type || 'audio/webm',
        name: file.name,
        sizeLabel: `${Math.max(1, Math.round(file.size / 1024))} KB`,
      },
    ])
    insertHtmlAtSelection(`
      <figure class="audio-note" contenteditable="false">
        <audio controls data-asset-id="${assetId}" src="${dataUrl}"></audio>
        <figcaption>${escapeHtml(file.name)}</figcaption>
      </figure>
    `)
  }

  const startAudioRecording = async () => {
    if (isRecordingAudio) return
    if (!navigator.mediaDevices?.getUserMedia) {
      setSaveLabel('Audio recording is not available here')
      return
    }

    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true })
      const mimeType = MediaRecorder.isTypeSupported('audio/webm')
        ? 'audio/webm'
        : MediaRecorder.isTypeSupported('audio/mp4')
          ? 'audio/mp4'
          : ''
      const recorder = new MediaRecorder(stream, mimeType ? { mimeType } : undefined)
      recordingChunksRef.current = []
      recorder.ondataavailable = (event) => {
        if (event.data.size > 0) {
          recordingChunksRef.current.push(event.data)
        }
      }
      recorder.onstop = async () => {
        const blob = new Blob(recordingChunksRef.current, {
          type: recorder.mimeType || 'audio/webm',
        })
        stream.getTracks().forEach((track) => track.stop())
        mediaRecorderRef.current = null
        recordingChunksRef.current = []
        setIsRecordingAudio(false)
        if (blob.size > 0) {
          await insertAudioNote(blob)
        }
      }
      mediaRecorderRef.current = recorder
      recorder.start()
      setIsRecordingAudio(true)
      setSaveLabel('Recording audio...')
    } catch {
      setSaveLabel('Microphone permission denied')
    }
  }

  const stopAudioRecording = () => {
    mediaRecorderRef.current?.stop()
  }

  const stopSpeechTranscription = () => {
    speechRecognitionRef.current?.stop()
  }

  const startSpeechTranscription = () => {
    if (isTranscribing) {
      stopSpeechTranscription()
      return
    }

    const SpeechRecognitionApi = window.SpeechRecognition ?? window.webkitSpeechRecognition
    if (!SpeechRecognitionApi) {
      setSaveLabel('Speech transcription is not available here')
      return
    }

    try {
      const recognition = new SpeechRecognitionApi()
      speechRecognitionRef.current = recognition
      speechTranscriptRef.current = ''
      recognition.continuous = true
      recognition.interimResults = true
      recognition.lang = 'en-US'
      recognition.onresult = (event) => {
        const transcript = Array.from(
          { length: event.results.length },
          (_, index) => event.results[index][0]?.transcript ?? '',
        )
          .join(' ')
          .replace(/\s+/g, ' ')
          .trim()
        speechTranscriptRef.current = transcript
        setSaveLabel(
          transcript
            ? `Transcribing: ${transcript.slice(0, 48)}${transcript.length > 48 ? '...' : ''}`
            : 'Listening for speech...',
        )
      }
      recognition.onerror = (event) => {
        speechRecognitionRef.current = null
        setIsTranscribing(false)
        setSaveLabel(event.error === 'not-allowed' ? 'Microphone permission denied' : 'Speech transcription stopped')
      }
      recognition.onend = () => {
        const transcript = speechTranscriptRef.current
        speechRecognitionRef.current = null
        speechTranscriptRef.current = ''
        setIsTranscribing(false)
        if (!transcript) {
          setSaveLabel('No transcript captured')
          return
        }

        insertTranscriptIntoPage(transcript)
        setSaveLabel('Transcript inserted into the page')
      }
      recognition.start()
      setIsTranscribing(true)
      setSaveLabel('Listening for speech...')
    } catch {
      speechRecognitionRef.current = null
      setIsTranscribing(false)
      setSaveLabel('Speech transcription is not available here')
    }
  }

  const beginInkStroke = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (!page || activeTab !== 'Draw' || isCurrentSectionLocked) return
    const rect = event.currentTarget.getBoundingClientRect()
    const point = {
      x: event.clientX - rect.left,
      y: event.clientY - rect.top,
    }
    const stroke: InkStroke = {
      color: drawColor,
      id: createId(),
      points: [point],
      width: drawColor === '#ffe266' ? 14 : 3,
    }
    inkDrawingRef.current = stroke
    event.currentTarget.setPointerCapture(event.pointerId)
    updatePage({ inkStrokes: [...page.inkStrokes, stroke] })
  }

  const moveInkStroke = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (!page || !inkDrawingRef.current || activeTab !== 'Draw' || isCurrentSectionLocked) return
    const rect = event.currentTarget.getBoundingClientRect()
    const point = {
      x: event.clientX - rect.left,
      y: event.clientY - rect.top,
    }
    const nextStroke = {
      ...inkDrawingRef.current,
      points: [...inkDrawingRef.current.points, point],
    }
    inkDrawingRef.current = nextStroke
    updatePage({
      inkStrokes: page.inkStrokes.map((stroke) => (stroke.id === nextStroke.id ? nextStroke : stroke)),
    })
  }

  const endInkStroke = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (!inkDrawingRef.current) return
    event.currentTarget.releasePointerCapture(event.pointerId)
    inkDrawingRef.current = null
  }

  const clearInkStrokes = () => {
    if (!page) return
    updatePage({ inkStrokes: [] })
  }

  const handleImageSelection = async (files: File[]) => {
    if (files.length === 0) return
    const images = await Promise.all(
      files
        .filter((file) => file.type.startsWith('image/'))
        .map(async (file) => ({
          assetId: createId(),
          dataUrl: await readFileAsDataUrl(file),
          kind: 'image' as const,
          mimeType: file.type,
          name: file.name,
          sizeLabel: `${Math.max(1, Math.round(file.size / 1024))} KB`,
        })),
    )
    if (images.length === 0) return

    addAssetsToState(
      images.map((image) => ({
        createdAt: new Date().toISOString(),
        dataUrl: image.dataUrl,
        id: image.assetId,
        kind: image.kind,
        mimeType: image.mimeType,
        name: image.name,
        sizeLabel: image.sizeLabel,
      })),
    )

    const html = images
      .map(
        (image) => `
          <figure class="embedded-image" contenteditable="false">
            <img alt="${escapeAttribute(image.name)}" data-asset-id="${image.assetId}" src="${image.dataUrl}" />
            <figcaption>${escapeHtml(image.name)}</figcaption>
          </figure>
        `,
      )
      .join('')
    insertHtmlAtSelection(html)
    if (imageInputRef.current) imageInputRef.current.value = ''
  }

  const handleAttachmentSelection = async (files: File[]) => {
    if (files.length === 0) return
    const attachments = await Promise.all(
      files.map(async (file) => ({
        assetId: createId(),
        dataUrl: await readFileAsDataUrl(file),
        kind: 'file' as const,
        mimeType: file.type || 'application/octet-stream',
        name: file.name,
        sizeLabel: `${Math.max(1, Math.round(file.size / 1024))} KB`,
      })),
    )

    addAssetsToState(
      attachments.map((file) => ({
        createdAt: new Date().toISOString(),
        dataUrl: file.dataUrl,
        id: file.assetId,
        kind: file.kind,
        mimeType: file.mimeType,
        name: file.name,
        sizeLabel: file.sizeLabel,
      })),
    )

    const html = attachments
      .map(
        (file) => `
          <div class="attachment-card" contenteditable="false" data-asset-id="${file.assetId}" data-download-url="${file.dataUrl}" data-file-name="${escapeAttribute(file.name)}">
            <strong class="attachment-title">${escapeHtml(file.name)}</strong>
            <span class="attachment-meta">${escapeHtml(file.sizeLabel)}</span>
          </div>
        `,
      )
      .join('')
    insertHtmlAtSelection(html)
    if (attachmentInputRef.current) attachmentInputRef.current.value = ''
  }

  const handlePrintoutSelection = async (files: File[]) => {
    if (files.length === 0) return
    const printouts = await Promise.all(
      files
        .filter((file) => file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf'))
        .map(async (file) => ({
          assetId: createId(),
          dataUrl: await readFileAsDataUrl(file),
          kind: 'printout' as const,
          mimeType: file.type || 'application/pdf',
          name: file.name,
          sizeLabel: `${Math.max(1, Math.round(file.size / 1024))} KB`,
        })),
    )
    if (printouts.length === 0) return

    addAssetsToState(
      printouts.map((file) => ({
        createdAt: new Date().toISOString(),
        dataUrl: file.dataUrl,
        id: file.assetId,
        kind: file.kind,
        mimeType: file.mimeType,
        name: file.name,
        sizeLabel: file.sizeLabel,
      })),
    )

    const html = printouts
      .map(
        (file) => `
          <section class="printout-card" contenteditable="false" data-asset-id="${file.assetId}" data-download-url="${file.dataUrl}" data-file-name="${escapeAttribute(file.name)}">
            <div class="printout-preview-shell">
              <iframe class="printout-preview" data-asset-id="${file.assetId}" src="${file.dataUrl}" title="${escapeAttribute(file.name)}"></iframe>
            </div>
            <div class="printout-caption">${escapeHtml(file.name)} · ${escapeHtml(file.sizeLabel)}</div>
          </section>
        `,
      )
      .join('')
    insertHtmlAtSelection(html)
    if (printoutInputRef.current) printoutInputRef.current.value = ''
  }

  const handleEditorPaste = async (event: ReactClipboardEvent<HTMLDivElement>) => {
    const clipboard = event.clipboardData
    const pastedFiles = [...clipboard.files]
    if (pastedFiles.length > 0) {
      event.preventDefault()
      const images = pastedFiles.filter((file) => file.type.startsWith('image/'))
      if (images.length > 0) {
        await handleImageSelection(images)
        return
      }
    }

    const html = clipboard.getData('text/html')
    const text = clipboard.getData('text/plain')
    if (!html && !text) return

    event.preventDefault()
    if (html) {
      insertHtmlAtSelection(sanitizePastedHtml(html))
      return
    }

    insertTextAsHtml(text)
  }

  const handleEditorClick = (event: ReactMouseEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement
    const linkedPageId = target.closest<HTMLElement>('[data-page-id]')?.dataset.pageId
    if (linkedPageId) {
      event.preventDefault()
      openLinkedPage(linkedPageId)
      return
    }

    const assetCard = target.closest<HTMLElement>('.attachment-card, .printout-card')
    if (assetCard) {
      const assetId = assetCard.dataset.assetId
      const asset = assetId ? appState.meta.assets[assetId] : undefined
      const downloadUrl = assetCard.dataset.downloadUrl ?? asset?.dataUrl
      const fileName = assetCard.dataset.fileName ?? asset?.name ?? 'attachment'
      if (downloadUrl) {
        const anchor = document.createElement('a')
        anchor.href = downloadUrl
        anchor.download = fileName
        anchor.click()
      }
      return
    }

    window.setTimeout(() => {
      syncEditorContent()
    }, 0)
  }

  const handleEditorKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Tab') {
      event.preventDefault()
      runEditorCommand(event.shiftKey ? 'outdent' : 'indent')
      return
    }

    if (event.key === 'Enter') {
      const checklistContext = getChecklistContext()
      if (checklistContext) {
        event.preventDefault()
        const nextItem = createChecklistItemNode()
        checklistContext.item.insertAdjacentElement('afterend', nextItem)
        focusChecklistItem(nextItem)
        window.setTimeout(syncEditorContent, 0)
        return
      }
    }

    if (event.key === 'Backspace') {
      const checklistContext = getChecklistContext()
      if (checklistContext) {
        const labelText =
          checklistContext.item.querySelector('label')?.textContent?.replace(/\s+/g, ' ').trim() ?? ''
        if (!labelText) {
          event.preventDefault()
          const nextItem =
            (checklistContext.item.previousElementSibling as HTMLLIElement | null) ??
            (checklistContext.item.nextElementSibling as HTMLLIElement | null)
          checklistContext.item.remove()

          if (checklistContext.list.children.length === 0) {
            const paragraph = document.createElement('p')
            paragraph.append(document.createTextNode(''))
            checklistContext.list.insertAdjacentElement('afterend', paragraph)
            checklistContext.list.remove()
            focusEditor(false)
          } else if (nextItem) {
            focusChecklistItem(nextItem)
          }

          window.setTimeout(syncEditorContent, 0)
          return
        }
      }
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
      event.preventDefault()
      insertExternalLink()
    }
  }

  const createSection = (groupId: string, name: string, color: string) => {
    setAppState((current) => {
      const sectionId = createId()
      const pageId = createId()
      const now = new Date().toISOString()
      return {
        ...current,
        notebooks: current.notebooks.map((item) =>
          item.id === current.selectedNotebookId
            ? {
                ...item,
                sectionGroups: item.sectionGroups.map((group) =>
                  group.id === groupId
                    ? {
                        ...group,
                        sections: [
                          ...group.sections,
                          {
                            color,
                            id: sectionId,
                            name,
                            passwordHash: null,
                            passwordHint: '',
                            pages: [
                              {
                                accent: color,
                                children: [],
                                content: '<p>New section notes.</p>',
                                createdAt: now,
                                id: pageId,
                                inkStrokes: [],
                                isCollapsed: false,
                                snippet: buildSnippet('Untitled Page', '<p>New section notes.</p>', now),
                                tags: [],
                                task: null,
                                title: 'Untitled Page',
                                updatedAt: now,
                              },
                            ],
                          },
                        ],
                      }
                    : group,
                ),
              }
            : item,
        ),
        selectedSectionGroupId: groupId,
        selectedPageId: pageId,
        selectedSectionId: sectionId,
      }
    })
  }

  const promptCreateSection = (groupId?: string) => {
    const targetGroupId = groupId ?? sectionGroup?.id
    if (!targetGroupId) return
    const name = window.prompt('New section name', 'New Section')?.trim()
    if (!name) return
    createSection(targetGroupId, name, '#4c75b8')
  }

  const addSectionGroup = () => {
    setAppState((current) => {
      const notebook = current.notebooks.find((item) => item.id === current.selectedNotebookId)
      if (!notebook) return current
      const groupName = window
        .prompt('New section group name', `Section Group ${notebook.sectionGroups.length + 1}`)
        ?.trim()
      if (!groupName) return current
      const sectionId = createId()
      const pageId = createId()
      const groupId = createId()
      const now = new Date().toISOString()
      return {
        ...current,
        notebooks: current.notebooks.map((item) =>
          item.id === notebook.id
            ? {
                ...item,
                sectionGroups: [
                  ...item.sectionGroups,
                  {
                    id: groupId,
                    isCollapsed: false,
                    name: groupName,
                    sections: [
                      {
                        color: '#4c75b8',
                        id: sectionId,
                        name: 'New Section',
                        passwordHash: null,
                        passwordHint: '',
                        pages: [
                          {
                            accent: '#4c75b8',
                            children: [],
                            content: '<p>New section group starter note.</p>',
                            createdAt: now,
                            id: pageId,
                            inkStrokes: [],
                            isCollapsed: false,
                            snippet: buildSnippet('Untitled Page', '<p>New section group starter note.</p>', now),
                            tags: [],
                            task: null,
                            title: 'Untitled Page',
                            updatedAt: now,
                          },
                        ],
                      },
                    ],
                  },
                ],
              }
            : item,
        ),
        selectedPageId: pageId,
        selectedSectionGroupId: groupId,
        selectedSectionId: sectionId,
      }
    })
  }

  const renameSectionGroup = (groupId: string) => {
    setAppState((current) => {
      const notebook = current.notebooks.find((item) => item.id === current.selectedNotebookId)
      const group = notebook?.sectionGroups.find((item) => item.id === groupId)
      if (!group) return current
      const nextName = window.prompt('Rename section group', group.name)?.trim()
      if (!nextName || nextName === group.name) return current
      return {
        ...current,
        notebooks: current.notebooks.map((item) =>
          item.id === current.selectedNotebookId
            ? {
                ...item,
                sectionGroups: item.sectionGroups.map((entry) =>
                  entry.id === groupId ? { ...entry, name: nextName } : entry,
                ),
              }
            : item,
        ),
      }
    })
  }

  const renameSection = (groupId: string, sectionId: string) => {
    setAppState((current) => {
      const notebook = current.notebooks.find((item) => item.id === current.selectedNotebookId)
      const group = notebook?.sectionGroups.find((item) => item.id === groupId)
      const currentSection = group?.sections.find((item) => item.id === sectionId)
      if (!currentSection) return current
      const nextName = window.prompt('Rename section', currentSection.name)?.trim()
      if (!nextName || nextName === currentSection.name) return current
      return {
        ...current,
        notebooks: current.notebooks.map((item) =>
          item.id === current.selectedNotebookId
            ? {
                ...item,
                sectionGroups: item.sectionGroups.map((entry) =>
                  entry.id === groupId
                    ? {
                        ...entry,
                        sections: entry.sections.map((part) =>
                          part.id === sectionId ? { ...part, name: nextName } : part,
                        ),
                      }
                    : entry,
                ),
              }
            : item,
        ),
      }
    })
  }

  const renameNotebook = (notebookId: string) => {
    setAppState((current) => {
      const currentNotebook = current.notebooks.find((item) => item.id === notebookId)
      if (!currentNotebook) return current
      const nextName = window.prompt('Rename notebook', currentNotebook.name)?.trim()
      if (!nextName || nextName === currentNotebook.name) return current

      return {
        ...current,
        notebooks: current.notebooks.map((item) =>
          item.id === notebookId ? { ...item, name: nextName } : item,
        ),
      }
    })
  }

  const deleteNotebook = (notebookId: string) => {
    if (!canDeleteNotebook) {
      window.alert('Keep at least one notebook until notebook bootstrap parity is finished.')
      return
    }

    const currentNotebook = appState.notebooks.find((item) => item.id === notebookId)
    if (!currentNotebook) return
    if (!window.confirm(`Delete notebook "${currentNotebook.name}" and all of its sections and pages?`)) {
      return
    }

    setAppState((current) =>
      ensureSelection({
        ...current,
        notebooks: current.notebooks.filter((item) => item.id !== notebookId),
      }),
    )
  }

  const deleteSectionGroup = (groupId: string) => {
    if (!notebook) return
    if (!canDeleteSectionGroup) {
      window.alert('Keep at least one section group in this notebook for now.')
      return
    }

    const group = notebook.sectionGroups.find((item) => item.id === groupId)
    if (!group) return
    if (!window.confirm(`Delete section group "${group.name}" and everything inside it?`)) {
      return
    }

    setAppState((current) =>
      ensureSelection({
        ...current,
        notebooks: current.notebooks.map((item) =>
          item.id === current.selectedNotebookId
            ? {
                ...item,
                sectionGroups: item.sectionGroups.filter((entry) => entry.id !== groupId),
              }
            : item,
        ),
      }),
    )
  }

  const deleteSection = (groupId: string, sectionId: string) => {
    if (!sectionGroup) return
    if (!canDeleteSection) {
      window.alert('Keep at least one section in this section group for now.')
      return
    }

    const group = notebook?.sectionGroups.find((item) => item.id === groupId)
    const currentSection = group?.sections.find((item) => item.id === sectionId)
    if (!currentSection) return
    if (!window.confirm(`Delete section "${currentSection.name}" and all of its pages?`)) {
      return
    }

    setAppState((current) =>
      ensureSelection({
        ...current,
        notebooks: current.notebooks.map((item) =>
          item.id === current.selectedNotebookId
            ? {
                ...item,
                sectionGroups: item.sectionGroups.map((entry) =>
                  entry.id === groupId
                    ? {
                        ...entry,
                        sections: entry.sections.filter((part) => part.id !== sectionId),
                      }
                    : entry,
                ),
              }
            : item,
        ),
      }),
    )
  }

  void notebookSectionCount
  void notebookPageCount
  void protectSection
  void removeSectionProtection
  void setSectionGroupDropTarget
  void setSectionDropTarget
  void setSectionGroupInsideDropTarget
  void moveSectionGroup
  void moveSection
  void moveSectionToGroup
  void toggleSectionGroupCollapse
  void renameSectionGroup
  void renameSection
  void deleteSectionGroup
  void deleteSection

  const createNotebook = () => {
    setAppState((current) => {
      const notebookId = createId()
      const groupId = createId()
      const sectionId = createId()
      const pageId = createId()
      const now = new Date().toISOString()
      return {
        meta: {
          ...current.meta,
          recentPageIds: [pageId, ...current.meta.recentPageIds.filter((id) => id !== pageId)].slice(0, 8),
        },
        notebooks: [
          ...current.notebooks,
          {
            color: '#8b63c9',
            icon: 'book',
            id: notebookId,
            name: `Notebook ${current.notebooks.length + 1}`,
            sectionGroups: [
              {
                id: groupId,
                isCollapsed: false,
                name: 'Sections',
                sections: [
                  {
                    color: '#4c75b8',
                    id: sectionId,
                    name: 'New Section',
                    passwordHash: null,
                    passwordHint: '',
                    pages: [
                      {
                        accent: '#4c75b8',
                        children: [],
                        content: '<p>Start writing here.</p>',
                        createdAt: now,
                        id: pageId,
                        inkStrokes: [],
                        isCollapsed: false,
                        snippet: buildSnippet('Welcome', '<p>Start writing here.</p>', now),
                        tags: [],
                        task: null,
                        title: 'Welcome',
                        updatedAt: now,
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
        selectedNotebookId: notebookId,
        selectedSectionGroupId: groupId,
        selectedPageId: pageId,
        selectedSectionId: sectionId,
      }
    })
  }

  const addPage = () => {
    if (!section) return
    setAppState((current) => {
      const now = new Date().toISOString()
      const nextPage: Page = {
        accent: section.color,
        children: [],
        content: '<p>New note.</p>',
        createdAt: now,
        id: createId(),
        inkStrokes: [],
        isCollapsed: false,
        snippet: buildSnippet('Untitled Page', '<p>New note.</p>', now),
        tags: [],
        task: null,
        title: 'Untitled Page',
        updatedAt: now,
      }
      return {
        ...current,
        notebooks: current.notebooks.map((item) => ({
          ...item,
          sectionGroups: item.sectionGroups.map((group) => ({
            ...group,
            sections: group.sections.map((entry) =>
              entry.id === current.selectedSectionId
                ? { ...entry, pages: [nextPage, ...entry.pages] }
                : entry,
            ),
          })),
        })),
        selectedPageId: nextPage.id,
      }
    })
  }

  const addSubpage = () => {
    if (!section || !page) return
    setAppState((current) => {
      const now = new Date().toISOString()
      const nextPage: Page = {
        accent: page.accent,
        children: [],
        content: '<p>New subpage.</p>',
        createdAt: now,
        id: createId(),
        inkStrokes: [],
        isCollapsed: false,
        snippet: buildSnippet('Untitled Subpage', '<p>New subpage.</p>', now),
        tags: [],
        task: null,
        title: 'Untitled Subpage',
        updatedAt: now,
      }
      return {
        ...current,
        notebooks: current.notebooks.map((item) => ({
          ...item,
          sectionGroups: item.sectionGroups.map((group) => ({
            ...group,
            sections: group.sections.map((entry) =>
              entry.id === current.selectedSectionId
                ? {
                    ...entry,
                    pages: updateNestedPages(entry.pages, current.selectedPageId, (note) => ({
                      ...note,
                      isCollapsed: false,
                      children: [...note.children, nextPage],
                    })),
                  }
                : entry,
            ),
          })),
        })),
        selectedPageId: nextPage.id,
      }
    })
  }

  const promoteCurrentPage = () => {
    if (!section || !page || !canPromotePage) return

    setAppState((current) => {
      let moved = false
      const nextState = {
        ...current,
        notebooks: current.notebooks.map((item) => ({
          ...item,
          sectionGroups: item.sectionGroups.map((group) => ({
            ...group,
            sections: group.sections.map((entry) => {
              if (entry.id !== current.selectedSectionId) return entry

              const promoted = promotePageOneLevel(entry.pages, current.selectedPageId)
              moved = promoted.moved
              return moved ? { ...entry, pages: promoted.pages } : entry
            }),
          })),
        })),
      }

      return moved ? nextState : current
    })
  }

  const demoteCurrentPage = () => {
    if (!section || !page || !canDemotePage) return

    setAppState((current) => {
      let moved = false
      const nextState = {
        ...current,
        notebooks: current.notebooks.map((item) => ({
          ...item,
          sectionGroups: item.sectionGroups.map((group) => ({
            ...group,
            sections: group.sections.map((entry) => {
              if (entry.id !== current.selectedSectionId) return entry

              const demoted = demotePageOneLevel(entry.pages, current.selectedPageId)
              moved = demoted.moved
              return moved ? { ...entry, pages: demoted.pages } : entry
            }),
          })),
        })),
      }

      return moved ? nextState : current
    })
  }

  const deleteCurrentPage = () => {
    if (!section || !page) return
    if (!canDeletePage) {
      window.alert('Keep at least one page in this section for now.')
      return
    }
    if (!window.confirm(`Delete page "${page.title}"?`)) return

    setAppState((current) => {
      const currentNotebook = current.notebooks.find((item) => item.id === current.selectedNotebookId)
      const currentGroup = currentNotebook?.sectionGroups.find(
        (item) => item.id === current.selectedSectionGroupId,
      )
      const currentSection = currentGroup?.sections.find((item) => item.id === current.selectedSectionId)
      if (!currentSection) return current

      const flattenedBeforeDelete = flattenPages(currentSection.pages, 0, true)
      const deletedIndex = flattenedBeforeDelete.findIndex((entry) => entry.page.id === current.selectedPageId)
      if (deletedIndex === -1) return current

      let nextSelectedPageId = current.selectedPageId
      let removed = false
      const nextState = {
        ...current,
        notebooks: current.notebooks.map((item) => ({
          ...item,
          sectionGroups: item.sectionGroups.map((group) => ({
            ...group,
            sections: group.sections.map((entry) => {
              if (entry.id !== current.selectedSectionId) return entry

              const result = removePageById(entry.pages, current.selectedPageId)
              if (!result.page) return entry

              removed = true
              const flattenedAfterDelete = flattenPages(result.pages, 0, true)
              const fallbackPage =
                flattenedAfterDelete[Math.min(deletedIndex, flattenedAfterDelete.length - 1)]?.page
              nextSelectedPageId = fallbackPage?.id ?? current.selectedPageId
              return { ...entry, pages: result.pages }
            }),
          })),
        })),
        selectedPageId: nextSelectedPageId,
      }

      return removed ? ensureSelection(nextState) : current
    })
  }

  const insertTemplate = (html: string) => {
    insertHtmlAtSelection(html)
  }

  const applyStylePreset = (html: string) => {
    insertTemplate(html)
    setIsStyleMenuOpen(false)
  }

  const applyFontFamily = (fontFamily: string) => {
    setSelectedFontFamily(fontFamily)
    runEditorCommand('fontName', fontFamily)
    setIsFontMenuOpen(false)
  }

  const applyFontSize = (fontSizeCommand: string, label: string) => {
    setSelectedFontSize(label)
    runEditorCommand('fontSize', fontSizeCommand)
    setIsFontSizeMenuOpen(false)
  }

  const pasteFromClipboard = async () => {
    if (navigator.clipboard?.read) {
      try {
        const items = await navigator.clipboard.read()
        for (const item of items) {
          if (item.types.includes('text/html')) {
            const blob = await item.getType('text/html')
            const html = await blob.text()
            insertHtmlAtSelection(sanitizePastedHtml(html))
            return
          }

          const imageType = item.types.find((type) => type.startsWith('image/'))
          if (imageType) {
            const blob = await item.getType(imageType)
            const extension = imageType.split('/')[1] ?? 'png'
            const image = new File([blob], `clipboard-image.${extension}`, { type: imageType })
            await handleImageSelection([image])
            return
          }

          if (item.types.includes('text/plain')) {
            const blob = await item.getType('text/plain')
            const text = await blob.text()
            if (text.trim()) {
              insertTextAsHtml(text)
              return
            }
          }
        }
      } catch {
        setSaveLabel('Clipboard permission blocked, using text-only paste')
      }
    }

    if (!navigator.clipboard?.readText) {
      setSaveLabel('Clipboard paste is not available here')
      return
    }

    const text = await navigator.clipboard.readText()
    if (!text.trim()) return
    insertTextAsHtml(text)
  }

  shortcutActionsRef.current = {
    addPage,
    createNotebook,
    demoteCurrentPage,
    openNotebook,
    promoteCurrentPage,
    saveNow,
    saveNotebookAs,
  }

  useEffect(() => {
    const handleShortcut = (event: KeyboardEvent) => {
      const target = event.target
      const isEditable =
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        (target instanceof HTMLElement && target.isContentEditable)

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault()
        void shortcutActionsRef.current.saveNow()
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
        shortcutActionsRef.current.addPage()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 'n') {
        event.preventDefault()
        shortcutActionsRef.current.createNotebook()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'o') {
        event.preventDefault()
        shortcutActionsRef.current.openNotebook()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 's') {
        event.preventDefault()
        void shortcutActionsRef.current.saveNotebookAs()
        return
      }

      if (isEditable) return

      if ((event.ctrlKey || event.metaKey) && event.key === 'ArrowUp' && canPromotePage) {
        event.preventDefault()
        shortcutActionsRef.current.promoteCurrentPage()
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key === 'ArrowDown' && canDemotePage) {
        event.preventDefault()
        shortcutActionsRef.current.demoteCurrentPage()
      }
    }

    window.addEventListener('keydown', handleShortcut)
    return () => {
      window.removeEventListener('keydown', handleShortcut)
    }
  }, [canDemotePage, canPromotePage])

  const renderNotebookIcon = (item: Notebook) => {
    if (item.icon === 'folder') {
      return <FolderIcon className="notebook-glyph-svg" color={item.color} />
    }
    if (item.icon === 'person') {
      return <PersonIcon className="notebook-glyph-svg" color={item.color} />
    }
    return <SectionBookIcon className="notebook-glyph-svg" color={item.color} />
  }

  const getSnippetParts = (entry: Page) => {
    const value =
      typeof entry.snippet === 'string' && entry.snippet.trim()
        ? entry.snippet
        : buildSnippet(entry.title, entry.content, entry.updatedAt)
    const [first, ...rest] = value.split('\n')
    return {
      first: first || formatPageDate(entry.updatedAt),
      second: rest.join(' ') || entry.title,
    }
  }

  const renderRibbon = () => {
    if (activeTab === 'File') {
      return (
        <section className="ribbon">
          <div className="ribbon-cluster styles">
            <button onClick={openNotebook} type="button">
              <FolderIcon size={26} />
              <span>Open Notebook</span>
            </button>
            <button onClick={() => void saveNotebookAs()} type="button">
              <SaveIcon size={26} />
              <span>Save Notebook As</span>
            </button>
            <button onClick={() => void saveNow()} type="button">
              <SaveIcon size={26} />
              <span>Save Notebook</span>
            </button>
            <button onClick={addPage} type="button">
              <EditIcon size={26} />
              <span>New Page</span>
            </button>
            <button onClick={addSubpage} type="button">
              <ListLinesIcon size={26} />
              <span>New Subpage</span>
            </button>
            <button onClick={() => promptCreateSection()} type="button">
              <SectionBookIcon size={26} />
              <span>New Section</span>
            </button>
            <button onClick={addSectionGroup} type="button">
              <FolderIcon size={26} />
              <span>New Section Group</span>
            </button>
            <button onClick={createNotebook} type="button">
              <NotebookStackIcon size={26} />
              <span>New Notebook</span>
            </button>
          </div>
          {recentNotebookEntries.length > 0 ? (
            <div className="ribbon-cluster styles">
              {recentNotebookEntries.slice(0, 4).map((entry) => (
                <button key={entry.path} onClick={() => void loadNotebookFromPath(entry.path)} type="button">
                  <FolderIcon size={26} />
                  <span>{entry.name}</span>
                </button>
              ))}
            </div>
          ) : null}
        </section>
      )
    }

    if (activeTab === 'Insert') {
      return (
        <section className="ribbon">
          <div className="ribbon-cluster styles">
            <button onClick={addPage} type="button">
              <EditIcon size={26} />
              <span>Blank Page</span>
            </button>
            <button onClick={insertChecklist} type="button">
              <TagsIcon size={26} />
              <span>Checklist</span>
            </button>
            <button onClick={insertTable} type="button">
              <TableIcon size={26} />
              <span>Table</span>
            </button>
            <button onClick={insertExternalLink} type="button">
              <LinkIcon size={26} />
              <span>Link</span>
            </button>
            <button onClick={insertInternalPageLink} type="button">
              <ProjectIcon size={26} />
              <span>Page Link</span>
            </button>
            <button onClick={openImagePicker} type="button">
              <ImageIcon size={26} />
              <span>Picture</span>
            </button>
            <button onClick={openPrintoutPicker} type="button">
              <ShowIcon size={26} />
              <span>Printout</span>
            </button>
            <button onClick={() => void (isRecordingAudio ? stopAudioRecording() : startAudioRecording())} type="button">
              <SaveIcon size={26} />
              <span>{isRecordingAudio ? 'Stop Audio' : 'Audio Note'}</span>
            </button>
            <button onClick={openAttachmentPicker} type="button">
              <AttachmentIcon size={26} />
              <span>File</span>
            </button>
            <button onClick={applyPageTemplate} type="button">
              <InsertFormattingIcon size={26} />
              <span>Template</span>
            </button>
            <button
              onClick={() => insertTemplate(`<p>${new Date().toLocaleDateString()}</p>`)}
              type="button"
            >
              <FormatMotivationIcon size={26} />
              <span>Date</span>
            </button>
            <span className="ribbon-label">Insert</span>
          </div>
        </section>
      )
    }

    if (activeTab === 'Draw') {
      return (
        <section className="ribbon">
          <div className="ribbon-cluster styles">
            <button onClick={() => setDrawColor('#1a73d9')} type="button">
              <PenIcon size={26} />
              <span>Blue Ink</span>
            </button>
            <button onClick={() => setDrawColor('#232a35')} type="button">
              <PenIcon size={26} />
              <span>Black Ink</span>
            </button>
            <button onClick={() => setDrawColor('#ffe266')} type="button">
              <BrushIcon size={26} />
              <span>Highlighter</span>
            </button>
            <button onClick={clearInkStrokes} type="button">
              <CutIcon size={26} />
              <span>Clear Ink</span>
            </button>
            <span className="ribbon-label">Draw</span>
          </div>
        </section>
      )
    }

    if (activeTab === 'History') {
      return (
        <section className="ribbon">
          <div className="ribbon-cluster styles">
            <button onClick={() => document.execCommand('undo')} type="button">
              <UndoIcon size={26} />
              <span>Undo</span>
            </button>
            <button onClick={() => document.execCommand('redo')} type="button">
              <ChevronDownIcon size={26} />
              <span>Redo</span>
            </button>
            <button onClick={() => saveCurrentPageVersion()} type="button">
              <SaveIcon size={26} />
              <span>Save Version</span>
            </button>
            <button onClick={restoreSavedPageVersion} type="button">
              <ShowIcon size={26} />
              <span>Restore Version</span>
            </button>
            <button onClick={() => setQuery('')} type="button">
              <SearchIcon size={26} />
              <span>Clear Search</span>
            </button>
            <span className="ribbon-label">History</span>
          </div>
        </section>
      )
    }

    if (activeTab === 'Review') {
      return (
        <section className="ribbon">
          <div className="ribbon-cluster styles">
            <button onClick={addTagToCurrentPage} type="button">
              <TagsIcon size={26} />
              <span>Add Tag</span>
            </button>
            <button onClick={toggleCurrentTask} type="button">
              <BulletsIcon size={26} />
              <span>{page?.task ? 'Remove Task' : 'Follow Up'}</span>
            </button>
            <button disabled={!page?.task} onClick={setCurrentTaskDueDate} type="button">
              <ProjectIcon size={26} />
              <span>Set Due Date</span>
            </button>
            <button disabled={!page?.task} onClick={toggleCurrentTaskComplete} type="button">
              <ShowIcon size={26} />
              <span>{page?.task?.status === 'done' ? 'Mark Open' : 'Mark Done'}</span>
            </button>
            <span className="ribbon-label">Review</span>
          </div>
        </section>
      )
    }

    if (activeTab === 'View') {
      return (
        <section className="ribbon">
          <div className="ribbon-cluster styles">
            <button onClick={() => editorRef.current?.scrollIntoView({ behavior: 'smooth' })} type="button">
              <ShowIcon size={26} />
              <span>Focus Page</span>
            </button>
            <button onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })} type="button">
              <ListLinesIcon size={26} />
              <span>Top</span>
            </button>
            <button onClick={() => void saveNow()} type="button">
              <SaveIcon size={26} />
              <span>Refresh Save</span>
            </button>
            <span className="ribbon-label">View</span>
          </div>
        </section>
      )
    }

    return (
      <section className="ribbon ribbon-home">
        <div className="ribbon-cluster clipboard">
          <div className="ribbon-big">
            <button onClick={() => void pasteFromClipboard()} type="button">
              <PasteIcon className="ribbon-large-icon" size={30} />
              <span>Paste</span>
              <ChevronDownIcon size={12} />
            </button>
          </div>
          <div className="ribbon-stack">
            <button onClick={() => document.execCommand('cut')} type="button">
              <CutIcon size={16} />
              <span>Cut</span>
            </button>
            <button onClick={() => void copySelection()} type="button">
              <CopyIcon size={16} />
              <span>Copy</span>
            </button>
            <button onClick={() => runEditorCommand('removeFormat')} type="button">
              <BrushIcon size={16} />
              <span>Format Painter</span>
            </button>
          </div>
          <span className="ribbon-label">Clipboard</span>
        </div>

        <div className="ribbon-cluster font home-basic-text">
          <div className="ribbon-row-compact">
            <div className="picker-dropdown" ref={fontMenuRef}>
              <button className="picker" onClick={() => setIsFontMenuOpen((current) => !current)} type="button">
                <span>{selectedFontFamily}</span>
                <ChevronDownIcon size={12} />
              </button>
              {isFontMenuOpen ? (
                createPortal(
                  <div
                    className="picker-menu picker-menu-floating"
                    ref={fontMenuPanelRef}
                    style={getFloatingMenuStyle(fontMenuRef.current, 164)}
                  >
                    {fontFamilies.map((fontFamily) => (
                      <button key={fontFamily} onClick={() => applyFontFamily(fontFamily)} type="button">
                        {fontFamily}
                      </button>
                    ))}
                  </div>,
                  document.body,
                )
              ) : null}
            </div>
            <div className="picker-dropdown" ref={fontSizeMenuRef}>
              <button
                className="picker tiny"
                onClick={() => setIsFontSizeMenuOpen((current) => !current)}
                type="button"
              >
                <span>{selectedFontSize}</span>
                <ChevronDownIcon size={12} />
              </button>
              {isFontSizeMenuOpen ? (
                createPortal(
                  <div
                    className="picker-menu picker-menu-floating tiny"
                    ref={fontSizeMenuPanelRef}
                    style={getFloatingMenuStyle(fontSizeMenuRef.current, 58)}
                  >
                    {fontSizes.map((fontSize) => (
                      <button
                        key={fontSize.command}
                        onClick={() => applyFontSize(fontSize.command, fontSize.label)}
                        type="button"
                      >
                        {fontSize.label}
                      </button>
                    ))}
                  </div>,
                  document.body,
                )
              ) : null}
            </div>
            <button onClick={() => runEditorCommand('fontSize', '4')} type="button">
              <TextSizeUpIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('fontSize', '2')} type="button">
              <TextSizeDownIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('foreColor', '#7e42b3')} type="button">
              <PenIcon size={16} />
            </button>
            <button onClick={insertExternalLink} type="button">
              <LinkIcon size={16} />
            </button>
          </div>
          <div className="ribbon-row-compact">
            <button className="strong" onClick={() => runEditorCommand('bold')} type="button">
              <BoldIcon size={16} />
            </button>
            <button className="strong" onClick={() => runEditorCommand('italic')} type="button">
              <ItalicIcon size={16} />
            </button>
            <button className="strong" onClick={() => runEditorCommand('underline')} type="button">
              <UnderlineIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('superscript')} type="button">
              <TextSizeUpIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('hiliteColor', '#fff59d')} type="button">
              <PenIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('foreColor', '#d83b01')} type="button">
              <BrushIcon size={16} />
            </button>
              <button onClick={insertChecklist} type="button">
                <TagsIcon size={16} />
              </button>
            </div>
          <span className="ribbon-label">Font</span>
        </div>

        <div className="ribbon-cluster paragraph home-basic-text">
          <div className="ribbon-row-compact">
            <button onClick={() => runEditorCommand('insertUnorderedList')} type="button">
              <BulletsIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('outdent')} type="button">
              <IndentIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('justifyLeft')} type="button">
              <AlignLeftIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('insertOrderedList')} type="button">
              <SortIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('selectAll')} type="button">
              <ShowIcon size={16} />
            </button>
            <button onClick={insertTable} type="button">
              <TableIcon size={16} />
            </button>
          </div>
          <div className="ribbon-row-compact">
            <button onClick={() => runEditorCommand('justifyLeft')} type="button">
              <AlignLeftIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('insertParagraph')} type="button">
              <ListLinesIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('justifyCenter')} type="button">
              <AlignLeftIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('indent')} type="button">
              <IndentIcon size={16} />
            </button>
            <button onClick={() => runEditorCommand('removeFormat')} type="button">
              <ShowIcon size={16} />
            </button>
            <button onClick={openImagePicker} type="button">
              <ImageIcon size={16} />
            </button>
          </div>
          <span className="ribbon-label">Paragraph</span>
        </div>

        <div className="ribbon-cluster home-styles">
          <div className="style-dropdown home-style-dropdown" ref={styleMenuRef}>
            <button className="home-style-card" onClick={() => setIsStyleMenuOpen((current) => !current)} type="button">
              <strong>Heading 1</strong>
              <span>Heading 2</span>
              <ChevronDownIcon size={12} />
            </button>
            {isStyleMenuOpen ? (
              createPortal(
                <div
                  className="style-dropdown-menu style-dropdown-menu-floating"
                  ref={styleMenuPanelRef}
                  style={getFloatingMenuStyle(styleMenuRef.current, 172)}
                >
                  {stylePresets.map((preset) => (
                    <button key={preset.id} onClick={() => applyStylePreset(preset.html)} type="button">
                      {preset.label}
                    </button>
                  ))}
                </div>,
                document.body,
              )
            ) : null}
          </div>
          <span className="ribbon-label">Styles</span>
        </div>

        <div className="ribbon-cluster home-tags">
          <button className="home-tag-card" onClick={toggleCurrentTask} type="button">
            <span className="home-tag-dot">v</span>
            <span>{page?.task ? 'To Do (Ctrl+1)' : 'To Do (Ctrl+1)'}</span>
          </button>
          <button
            className="home-tag-card"
            onClick={() => insertTemplate('<p><span style="color:#d68200;">Important</span></p>')}
            type="button"
          >
            <span className="home-tag-star">*</span>
            <span>Important (Ctrl+2)</span>
          </button>
          <span className="ribbon-label">Tags</span>
        </div>

        <div className="ribbon-cluster home-quick-actions">
          <button className="home-vertical-action" onClick={addTagToCurrentPage} type="button">
            <TagsIcon size={22} />
            <span>To Do Tag</span>
          </button>
          <button className="home-vertical-action" onClick={() => insertTemplate('<p><strong>Tagged:</strong> </p>')} type="button">
            <FormatMotivationIcon size={22} />
            <span>Find Outlook Tasks</span>
          </button>
          <button className="home-vertical-action" onClick={emailCurrentPage} type="button">
            <ProjectIcon size={22} />
            <span>Email Page</span>
          </button>
          <button className="home-vertical-action" onClick={applyPageTemplate} type="button">
            <TableIcon size={22} />
            <span>Meeting Details</span>
          </button>
          <button
            className="home-vertical-action"
            onClick={() => void (isRecordingAudio ? stopAudioRecording() : startAudioRecording())}
            type="button"
          >
            <SaveIcon size={22} />
            <span>Dictate</span>
          </button>
          <button className="home-vertical-action" onClick={startSpeechTranscription} type="button">
            <ShowIcon size={22} />
            <span>{isTranscribing ? 'Stop Transcribe' : 'Transcribe'}</span>
          </button>
          <button className="home-vertical-action" onClick={() => setIsCopilotOpen(true)} type="button">
            <InsertFormattingIcon size={22} />
            <span>Copilot</span>
          </button>
        </div>
      </section>
    )
  }

  const dragLabel =
    dragState?.type === 'notebook'
      ? 'Moving notebook'
      : dragState?.type === 'section-group'
      ? 'Moving section group'
      : dragState?.type === 'section'
        ? 'Moving section'
        : dragState?.type === 'page'
          ? 'Moving page'
          : ''
  const windowTitle = `${page?.title?.trim() || 'Untitled Page'} - ${notebook?.name ?? 'Notebook'} - OneNote`
  const saveStatusText = isDirty ? `${saveLabel} · Unsaved changes` : saveLabel
  const displayVersion = appInfo?.version ?? __APP_VERSION__
  const suggestedPrompts = [
    'Change this bulleted list into full sentences and paragraphs',
    'Draft a plan for a team offsite in Santa Fe',
    'Give me ideas for ways to improve productivity and better manage my time',
  ]

  return (
    <div className={`desktop-scene ${dragState ? 'drag-active' : ''}`}>
      <div className={`onenote-window ${dragState ? `dragging-${dragState.type}` : ''}`}>
        <header className="titlebar">
          <div className="titlebar-left">
            <div className="titlebar-brand">
              <div className="app-badge">
                <OneNoteLogoIcon size={20} />
              </div>
              <button className="quick-action icon-only" onClick={() => void saveNow()} type="button">
                <SaveIcon size={16} />
              </button>
              <button
                className="quick-action icon-only"
                onClick={() => runEditorCommand('undo')}
                type="button"
              >
                <UndoIcon size={16} />
              </button>
              <button className="quick-action icon-only small" type="button">
                <ChevronDownIcon size={12} />
              </button>
            </div>
            <div className="titlebar-workspace" title={windowTitle}>
              <span className="workspace-name">{notebook?.name ?? 'Dunder Mifflin offsite'}</span>
              <span className="workspace-dot" />
              <span className="workspace-pill">Non-Business</span>
              <span className="workspace-dot" />
              <span className="workspace-app">Microsoft OneNote</span>
              <ChevronDownIcon size={11} />
            </div>
          </div>
          <div className="titlebar-search" ref={titlebarSearchRef}>
            <span className="search-icon">
              <SearchIcon size={16} />
            </span>
            <input
              ref={searchInputRef}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Search"
              type="search"
              value={query}
            />
            {searchResults.length > 0 ? (
              createPortal(
                <div
                  className="search-results search-results-floating"
                  ref={searchResultsPanelRef}
                  style={getFloatingPanelStyle(titlebarSearchRef.current)}
                >
                  {searchResults.slice(0, 8).map((result) => (
                    <button
                      key={result.page.id}
                      className="search-result"
                      onClick={() => openSearchResult(result)}
                      type="button"
                    >
                      <strong>{result.isSubpage ? `${result.page.title} (Subpage)` : result.page.title}</strong>
                      <span>{result.notebookName} / {result.groupName} / {result.sectionName}</span>
                    </button>
                  ))}
                </div>,
                document.body,
              )
            ) : null}
          </div>
          <div className="titlebar-right">
            <button className="titlebar-chip" type="button">
              Sticky Notes
            </button>
            <button className="titlebar-share" type="button">
              Share
            </button>
            <div className="profile-chip">SJ</div>
            <div className="window-controls" aria-hidden="true">
              <button className="window-control" type="button">-</button>
              <button className="window-control" type="button">[]</button>
              <button className="window-control close" type="button">x</button>
            </div>
          </div>
        </header>

        <nav className="tab-row">
          {ribbonTabs.map((item) => (
            <button
              key={item}
              className={`tab-button ${item === activeTab ? 'active' : ''}`}
              onClick={() => setActiveTab(item)}
              type="button"
            >
              {item}
            </button>
          ))}
        </nav>

        {renderRibbon()}

        <main className="workspace">
          <aside className="notebooks-pane">
            <div className="pane-heading">
              <span className="heading-icon">
                <NotebookStackIcon size={18} />
              </span>
              <div className="pane-heading-copy">
                <span className="pane-kicker">NOTEBOOKS</span>
                <strong>My Notebooks</strong>
              </div>
              <button className="pane-heading-action" onClick={createNotebook} type="button">
                <NotebookStackIcon size={14} />
                <span>New</span>
              </button>
              <span className="pane-count">{appState.notebooks.length}</span>
            </div>
            <div className="notebook-tree">
              {appState.notebooks.map((item) => (
                <div key={item.id} className="notebook-tree-group">
                  <div className="notebook-row">
                    <button
                      className={`notebook-item ${item.id === notebook?.id ? 'active' : ''} ${dragState?.type === 'notebook' && dragState.notebookId === item.id ? 'dragging' : ''} ${dropTarget?.type === 'notebook' && dropTarget.notebookId === item.id && dropTarget.position === 'before' ? 'drop-before' : ''} ${dropTarget?.type === 'notebook' && dropTarget.notebookId === item.id && dropTarget.position === 'after' ? 'drop-after' : ''}`}
                      onClick={() => {
                        if (consumeSuppressedClick()) return
                        selectNotebook(item.id)
                      }}
                      onPointerDown={(event) => beginDrag(event, { type: 'notebook', notebookId: item.id })}
                      onPointerEnter={allowDrop}
                      onPointerMove={(event) => setNotebookDropTarget(event, item.id)}
                      onPointerUp={() => {
                        if (dropTarget?.type === 'notebook' && dropTarget.notebookId === item.id) {
                          moveNotebook(item.id, dropTarget.position)
                        }
                      }}
                      onContextMenu={(event) => {
                        selectNotebook(item.id)
                        openContextMenu(event, [
                          { label: 'Rename notebook', onSelect: () => renameNotebook(item.id) },
                          { label: 'New section group', onSelect: addSectionGroup },
                          {
                            danger: true,
                            disabled: !canDeleteNotebook,
                            label: 'Delete notebook',
                            onSelect: () => deleteNotebook(item.id),
                          },
                        ])
                      }}
                      type="button"
                    >
                      <span className="notebook-glyph">{renderNotebookIcon(item)}</span>
                      <span>{item.name}</span>
                    </button>
                  </div>
                  <div className="notebook-sections">
                    {item.sectionGroups.flatMap((group) =>
                      group.sections.map((entry) => (
                        <button
                          key={entry.id}
                          className={`notebook-section-link ${item.id === notebook?.id && entry.id === section?.id ? 'active' : ''}`}
                          onClick={() => {
                            selectNotebook(item.id)
                            selectSection(group.id, entry.id)
                          }}
                          type="button"
                        >
                          <span className="section-color" style={{ backgroundColor: entry.color }} />
                          <span>{entry.name}</span>
                        </button>
                      )),
                    )}
                    <button
                      className="notebook-new-section"
                      onClick={() => promptCreateSection(item.sectionGroups[0]?.id)}
                      type="button"
                    >
                      + New section
                    </button>
                  </div>
                </div>
              ))}
            </div>
            <div className="pane-footer">
              <button
                className="settings-button"
                onClick={() => void runUpdateCheck('manual')}
                type="button"
              >
                <SettingsIcon size={16} />
                <span>{isCheckingForUpdates ? 'Checking...' : 'Check for updates'}</span>
              </button>
            </div>
          </aside>

          <aside className="pages-pane">
            <div className="pages-header">
              <div className="pages-heading-copy">
                <span className="pane-kicker">PAGES</span>
                <h2>{section?.name ?? 'Pages'}</h2>
                <span className="pages-context">
                  {sectionGroup?.name ?? 'Sections'} | {visiblePages.length}
                </span>
              </div>
              <div className="pages-actions">
                <div className="pages-actions-primary">
                  <button
                    className="pages-action-pill primary"
                    disabled={isCurrentSectionLocked}
                    onClick={addPage}
                    type="button"
                  >
                    <ListLinesIcon size={15} />
                    <span>Add page</span>
                  </button>
                  <button
                    className="pages-action-pill"
                    disabled={isCurrentSectionLocked}
                    onClick={addSubpage}
                    type="button"
                  >
                    <EditIcon size={15} />
                    <span>Subpage</span>
                  </button>
                </div>
                <div className="pages-actions-secondary">
                  <select
                    className="pages-sort-picker"
                    onChange={(event) => setPageSortMode(event.target.value as PageSortMode)}
                    value={pageSortMode}
                  >
                    {Object.entries(pageSortModeLabels).map(([value, label]) => (
                      <option key={value} value={value}>
                        {label}
                      </option>
                    ))}
                  </select>
                  <button
                    disabled={isCurrentSectionLocked || !canPromotePage}
                    onClick={promoteCurrentPage}
                    title="Promote page"
                    type="button"
                  >
                    <IndentIcon size={16} />
                  </button>
                  <button
                    disabled={isCurrentSectionLocked || !canDemotePage}
                    onClick={demoteCurrentPage}
                    title="Make subpage"
                    type="button"
                  >
                    <SubpageIcon size={16} />
                  </button>
                  <button
                    className="danger"
                    disabled={isCurrentSectionLocked || !canDeletePage}
                    onClick={deleteCurrentPage}
                    title="Delete page"
                    type="button"
                  >
                    <DeleteIcon size={16} />
                  </button>
                </div>
              </div>
            </div>
            {isCurrentSectionLocked ? (
              <div className="section-lock-panel">
                <strong>{section?.name ?? 'Protected section'}</strong>
                <span>{section?.passwordHint ? `Hint: ${section.passwordHint}` : 'This section is password protected.'}</span>
                <button onClick={() => void unlockSection(section?.id ?? '')} type="button">
                  Unlock Section
                </button>
              </div>
            ) : (
            <div className="page-cards">
              {visiblePages.map((entry) => {
                const snippet = getSnippetParts(entry.page)
                const hasChildren = entry.page.children.length > 0
                const hasSelectedChild = hasChildren && hasChildPageSelected(entry.page, page?.id ?? '')
                const pageLocation = section ? findPageLocation(section.pages, entry.page.id) : undefined
                const canPromoteEntry = Boolean(pageLocation?.parentId)
                const canDemoteEntry = Boolean(pageLocation && pageLocation.index > 0)
                return (
                  <div key={entry.page.id} className={`page-row ${entry.depth > 0 ? 'subpage-row' : ''}`} style={{ marginLeft: `${entry.depth * 24}px` }}>
                    {hasChildren ? (
                      <button
                        aria-label={entry.page.isCollapsed ? 'Expand subpages' : 'Collapse subpages'}
                        className={`page-disclosure ${entry.page.isCollapsed ? 'collapsed' : ''}`}
                        onClick={() => togglePageCollapse(entry.page.id)}
                        type="button"
                      >
                        <ChevronDownIcon size={12} />
                      </button>
                    ) : (
                      <span className="page-disclosure-placeholder" />
                    )}
                    <button
                      className={`page-card ${entry.page.id === page?.id ? 'active' : ''} ${entry.depth > 0 ? 'subpage-card' : ''} ${hasSelectedChild ? 'has-selected-child' : ''} ${dragState?.type === 'page' && dragState.pageId === entry.page.id ? 'dragging' : ''} ${dropTarget?.type === 'page' && dropTarget.pageId === entry.page.id && dropTarget.position === 'before' ? 'drop-before' : ''} ${dropTarget?.type === 'page' && dropTarget.pageId === entry.page.id && dropTarget.position === 'after' ? 'drop-after' : ''}`}
                      onPointerDown={(event) => beginDrag(event, { type: 'page', pageId: entry.page.id })}
                      onPointerEnter={allowDrop}
                      onPointerMove={(event) => setPageDropTarget(event, entry.page.id)}
                      onPointerUp={() => {
                        if (dropTarget?.type === 'page' && dropTarget.pageId === entry.page.id) {
                          movePage(entry.page.id, dropTarget.position)
                        }
                      }}
                      onClick={() => {
                        if (consumeSuppressedClick()) return
                        selectPage(entry.page.id)
                      }}
                      onContextMenu={(event) => {
                        selectPage(entry.page.id)
                        openContextMenu(event, [
                          { label: 'Rename page', onSelect: () => renamePage(entry.page.id) },
                          { label: 'New page', onSelect: addPage },
                          { label: 'New subpage', onSelect: addSubpage },
                          {
                            disabled: !canPromoteEntry,
                            label: 'Promote page',
                            onSelect: () => {
                              selectPage(entry.page.id)
                              window.setTimeout(promoteCurrentPage, 0)
                            },
                          },
                          {
                            disabled: !canDemoteEntry,
                            label: 'Make subpage',
                            onSelect: () => {
                              selectPage(entry.page.id)
                              window.setTimeout(demoteCurrentPage, 0)
                            },
                          },
                          { label: 'Save page version', onSelect: () => saveCurrentPageVersion(entry.page) },
                          {
                            danger: true,
                            disabled: !canDeletePage,
                            label: 'Delete page',
                            onSelect: () => {
                              selectPage(entry.page.id)
                              window.setTimeout(deleteCurrentPage, 0)
                            },
                          },
                        ])
                      }}
                      type="button"
                    >
                      <span className="page-accent" style={{ backgroundColor: entry.page.accent }} />
                      <div className="page-card-body">
                        <strong>{entry.depth > 0 ? `- ${entry.page.title}` : entry.page.title}</strong>
                        <span className="page-date">
                          {snippet.first}
                        </span>
                        <span className="page-snippet">{snippet.second}</span>
                      </div>
                    </button>
                  </div>
                )
              })}
            </div>
            )}
            {recentPages.length > 0 ? (
              <div className="recent-pages-panel">
                <div className="recent-pages-heading">
                  <span>Recent</span>
                </div>
                <div className="recent-pages-list">
                  {recentPages.slice(0, 4).map((item) => (
                    <button
                      key={item.page.id}
                      className={`recent-page-item ${item.page.id === page?.id ? 'active' : ''}`}
                      onClick={() =>
                        openSearchResult({
                          ...item,
                          isSubpage: false,
                        })
                      }
                      type="button"
                    >
                      <strong>{item.page.title}</strong>
                      <span>{item.sectionName}</span>
                    </button>
                  ))}
                </div>
              </div>
            ) : null}
          </aside>

          <section className={`note-pane ${isCopilotOpen ? 'copilot-open' : 'copilot-closed'}`}>
            <div className="note-document">
              <div className="note-header">
                <div className="note-header-main">
                  <div className="note-header-copy">
                    <input
                      className="note-title"
                      disabled={isCurrentSectionLocked}
                      onChange={(event) => updatePage({ title: event.target.value })}
                      type="text"
                      value={page?.title ?? ''}
                    />
                    <div className="note-date-row">
                      <span>{page ? formatPageDate(page.createdAt) : ''}</span>
                      <span>{page ? formatPageTime(page.createdAt) : ''}</span>
                    </div>
                  </div>
                </div>
                <div className="note-toolbar-inline">
                  <button disabled={isCurrentSectionLocked} onClick={addTagToCurrentPage} type="button">
                    Tag
                  </button>
                  <button disabled={isCurrentSectionLocked} onClick={toggleCurrentTask} type="button">
                    {page?.task ? 'Task On' : 'Task'}
                  </button>
                  <button
                    disabled={isCurrentSectionLocked}
                    onClick={() => void (isRecordingAudio ? stopAudioRecording() : startAudioRecording())}
                    type="button"
                  >
                    {isRecordingAudio ? 'Stop Audio' : 'Audio'}
                  </button>
                  <button disabled={isCurrentSectionLocked} onClick={() => saveCurrentPageVersion()} type="button">
                    Version
                  </button>
                </div>
              </div>
              <div className="note-canvas-shell">
                {isCurrentSectionLocked ? (
                  <div className="section-lock-screen">
                    <strong>{section?.name ?? 'Protected section'}</strong>
                    <p>{section?.passwordHint ? `Hint: ${section.passwordHint}` : 'Unlock this section to view and edit its notes.'}</p>
                    <button onClick={() => void unlockSection(section?.id ?? '')} type="button">
                      Unlock Section
                    </button>
                  </div>
                ) : (
                  <>
                    {page && (activeTab === 'Draw' || page.inkStrokes.length > 0) ? (
                      <div className="ink-board-shell">
                        <div className="ink-board-header">
                          <strong>Ink Layer</strong>
                          <span>{drawColor === '#ffe266' ? 'Highlighter' : `Pen ${drawColor}`}</span>
                        </div>
                        <svg
                          className="ink-board"
                          onPointerDown={beginInkStroke}
                          onPointerMove={moveInkStroke}
                          onPointerUp={endInkStroke}
                          onPointerLeave={endInkStroke}
                          ref={drawSurfaceRef}
                          viewBox="0 0 900 260"
                        >
                          {page.inkStrokes.map((stroke) => (
                            <polyline
                              key={stroke.id}
                              fill="none"
                              points={stroke.points.map((point) => `${point.x},${point.y}`).join(' ')}
                              stroke={stroke.color}
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeOpacity={stroke.color === '#ffe266' ? 0.5 : 1}
                              strokeWidth={stroke.width}
                            />
                          ))}
                        </svg>
                      </div>
                    ) : null}
                    <div
                      className="editor-canvas"
                      contentEditable
                      onBlur={syncEditorContent}
                      onClick={handleEditorClick}
                      onInput={syncEditorContent}
                      onKeyDown={handleEditorKeyDown}
                      onPaste={(event) => void handleEditorPaste(event)}
                      ref={editorRef}
                      suppressContentEditableWarning
                    />
                  </>
                )}
              </div>
              <div className="status-strip">
                <span>{saveStatusText}</span>
                <span>
                  {appInfo?.name ?? 'OnePlace'} v{displayVersion} | {searchScopeLabels[searchScope]}
                </span>
              </div>
            </div>
            <aside className="copilot-pane" hidden={!isCopilotOpen}>
              <div className="copilot-pane-toolbar">
                <span className="copilot-pane-icon">=</span>
                <span className="copilot-pane-status">o</span>
                <span className="copilot-pane-spacer" />
                <button type="button">...</button>
                <button onClick={() => setIsCopilotOpen(false)} type="button">x</button>
              </div>
              <div className="copilot-pane-body">
                <div className="copilot-orb" />
                <h3>Try 'Organize my Notes'</h3>
                <div className="copilot-compose">
                  <span>Message Copilot</span>
                  <div className="copilot-compose-actions">
                    <button type="button">+</button>
                    <button type="button">..</button>
                    <button type="button">o</button>
                    <button type="button">))</button>
                  </div>
                </div>
                <div className="copilot-prompts">
                  {suggestedPrompts.map((item) => (
                    <button key={item} className="copilot-prompt-card" type="button">
                      {item}
                    </button>
                  ))}
                </div>
                <button className="copilot-show-more" type="button">
                  Show more
                </button>
              </div>
            </aside>
          </section>
        </main>
      </div>
      <input
        accept="image/*"
        hidden
        onChange={(event: ReactChangeEvent<HTMLInputElement>) =>
          void handleImageSelection(Array.from(event.target.files ?? []))
        }
        ref={imageInputRef}
        type="file"
      />
      <input
        hidden
        onChange={(event: ReactChangeEvent<HTMLInputElement>) =>
          void handleAttachmentSelection(Array.from(event.target.files ?? []))
        }
        ref={attachmentInputRef}
        type="file"
      />
      <input
        accept=".pdf,application/pdf"
        hidden
        onChange={(event: ReactChangeEvent<HTMLInputElement>) =>
          void handlePrintoutSelection(Array.from(event.target.files ?? []))
        }
        ref={printoutInputRef}
        type="file"
      />
      {contextMenu ? (
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
                setContextMenu(null)
                item.onSelect()
              }}
              type="button"
            >
              {item.label}
            </button>
          ))}
        </div>
      ) : null}
      {dragState && dragPosition ? (
        <div
          className="drag-badge"
          style={{ left: `${dragPosition.x + 18}px`, top: `${dragPosition.y + 18}px` }}
        >
          {dragLabel}
        </div>
      ) : null}
    </div>
  )
}

export default App
