import type { Page } from './appTypes'

export const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

export const formatDate = (value: string) =>
  new Intl.DateTimeFormat('en-CA', {
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    month: 'short',
  }).format(new Date(value))

export const formatPageDate = (value: string) =>
  new Intl.DateTimeFormat('en-CA', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  }).format(new Date(value))

export const formatPageTime = (value: string) =>
  new Intl.DateTimeFormat('en-CA', {
    hour: 'numeric',
    minute: '2-digit',
  }).format(new Date(value))

export const formatDueDate = (value: string) =>
  new Intl.DateTimeFormat('en-CA', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
  }).format(new Date(value))

export const createId = () => crypto.randomUUID()

export const extractSnippetText = (content: string) => {
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

export const extractExportText = (content: string) => {
  if (typeof DOMParser === 'undefined') {
    return content
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|h1|h2|h3|li|tr|blockquote|pre|section)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim()
  }

  const doc = new DOMParser().parseFromString(content, 'text/html')

  for (const note of doc.querySelectorAll('.audio-note')) {
    note.textContent = 'Audio note'
  }

  for (const card of doc.querySelectorAll('.attachment-card')) {
    const title = card.querySelector('.attachment-title')?.textContent?.trim()
    const meta = card.querySelector('.attachment-meta')?.textContent?.trim()
    card.textContent = [title, meta].filter(Boolean).join(' - ')
  }

  for (const card of doc.querySelectorAll('.printout-card')) {
    const caption = card.querySelector('.printout-caption')?.textContent?.trim()
    card.textContent = caption || 'Printout'
  }

  const lines: string[] = []
  const blockTags = new Set([
    'P',
    'DIV',
    'SECTION',
    'ARTICLE',
    'H1',
    'H2',
    'H3',
    'H4',
    'H5',
    'H6',
    'BLOCKQUOTE',
    'PRE',
    'UL',
    'OL',
  ])

  const serializeNode = (node: Node): string[] => {
    if (node instanceof Text) {
      const value = node.textContent?.replace(/\s+/g, ' ').trim() ?? ''
      return value ? [value] : []
    }

    if (!(node instanceof HTMLElement)) return []
    if (node.tagName === 'BR') return ['\n']

    if (node.tagName === 'LI') {
      const contentText = serializeChildren(node).join(' ').replace(/\s+/g, ' ').trim()
      return contentText ? [`- ${contentText}`] : []
    }

    if (node.tagName === 'TR') {
      const row = [...node.querySelectorAll(':scope > th, :scope > td')]
        .map((cell) => cell.textContent?.replace(/\s+/g, ' ').trim() ?? '')
        .filter(Boolean)
        .join('\t')
      return row ? [row] : []
    }

    if (node.tagName === 'TABLE') {
      const rows = [...node.querySelectorAll('tr')]
        .flatMap((row) => serializeNode(row))
        .filter(Boolean)
      return rows.length > 0 ? [rows.join('\n')] : []
    }

    const childLines = serializeChildren(node)
    const joined = childLines.join(node.tagName === 'PRE' ? '\n' : ' ').replace(/[ \t]+\n/g, '\n')
    const normalized = joined.replace(/\n{3,}/g, '\n\n').replace(/[ \t]{2,}/g, ' ').trim()
    if (!normalized) return []
    return blockTags.has(node.tagName) ? [normalized, ''] : [normalized]
  }

  const serializeChildren = (element: HTMLElement) => [...element.childNodes].flatMap((child) => serializeNode(child))

  for (const child of [...doc.body.childNodes]) {
    lines.push(...serializeNode(child))
  }

  return lines.join('\n').replace(/\n{3,}/g, '\n\n').replace(/[ \t]+\n/g, '\n').trim()
}

export const buildSnippet = (title: string, content: string, timestamp = new Date().toISOString()) => {
  const plain = extractSnippetText(content)
  return `${formatPageDate(timestamp)}\n${plain || title}`.slice(0, 120)
}

export const buildSearchSnippet = (text: string, needle: string) => {
  const compact = text.replace(/\s+/g, ' ').trim()
  if (!compact) return 'No text preview available.'
  const index = compact.toLowerCase().indexOf(needle)
  if (index === -1) return compact.slice(0, 140)
  const start = Math.max(0, index - 36)
  const end = Math.min(compact.length, index + needle.length + 72)
  const prefix = start > 0 ? '...' : ''
  const suffix = end < compact.length ? '...' : ''
  return `${prefix}${compact.slice(start, end)}${suffix}`
}

export const formatElapsedTime = (totalSeconds: number) => {
  const minutes = Math.floor(totalSeconds / 60)
  const seconds = totalSeconds % 60
  return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`
}

export const createPage = (
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
