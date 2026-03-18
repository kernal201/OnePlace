import { useState, type Dispatch, type SetStateAction } from 'react'
import { ONEPLACE_ONENOTE_CLIENT_ID } from '../../app/appConfig'
import {
  accentPalette,
  buildSnippet,
  createSectionGroup,
  ensureSelection,
  sanitizePastedHtml,
} from '../../app/appModel'
import type { AppAsset, AppState, Notebook, Page, Section, SectionGroup } from '../../app/appModel'

const ONE_NOTE_IMPORT_ASSET_PREFIX = 'onenote-asset:'
const ONE_NOTE_GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'
const ONE_NOTE_AUTH_ROOT = 'https://login.microsoftonline.com/common/oauth2/v2.0'
const ONE_NOTE_SCOPES = ['openid', 'profile', 'offline_access', 'Notes.Read']
const ONE_NOTE_PAGE_FETCH_CONCURRENCY = 3

type DeviceCodeResponse = {
  device_code: string
  expires_in: number
  interval: number
  message?: string
  user_code: string
  verification_uri: string
  verification_uri_complete?: string
}

type TokenResponse = {
  access_token: string
}

type GraphCollectionResponse<T> = {
  '@odata.nextLink'?: string
  value?: T[]
}

type GraphNotebook = {
  displayName: string
  id: string
}

type GraphSectionGroup = {
  displayName: string
  id: string
}

type GraphSection = {
  displayName: string
  id: string
}

type GraphPage = {
  createdDateTime?: string
  id: string
  lastModifiedDateTime?: string
  level?: number
  order?: number
  title?: string
}

type ImportedPageNode = {
  assets: AppAsset[]
  level: number
  page: Page
}

const sleep = (ms: number) =>
  new Promise<void>((resolve) => {
    window.setTimeout(resolve, ms)
  })

const blobToDataUrl = (blob: Blob) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '')
    reader.onerror = () => reject(reader.error ?? new Error('Unable to read file data.'))
    reader.readAsDataURL(blob)
  })

const formatSizeLabel = (size: number) => {
  if (size >= 1024 * 1024) {
    return `${Math.max(1, Math.round((size / (1024 * 1024)) * 10) / 10)} MB`
  }

  return `${Math.max(1, Math.round(size / 1024))} KB`
}

const stripHtml = (value: string) =>
  value
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

const getErrorMessage = (error: unknown) => (error instanceof Error ? error.message : String(error))

const normalizeOneNoteImportError = (error: unknown) => {
  const message = getErrorMessage(error)

  if (/Selected user account does not exist in tenant/i.test(message)) {
    return 'The built-in OnePlace Microsoft app is registered against the wrong tenant or account audience. Update the built-in Client ID to an app registration that supports the accounts you want to import from.'
  }

  if (/AADSTS700016/i.test(message) || /application .* was not found/i.test(message)) {
    return 'The built-in OnePlace Microsoft app is not configured correctly. Update the built-in Microsoft Client ID before using OneNote import.'
  }

  return message
}

const parseJsonSafely = (value: string) => {
  try {
    return JSON.parse(value) as Record<string, unknown>
  } catch {
    return null
  }
}

const buildResponseError = async (response: Response) => {
  const raw = await response.text()
  const parsed = parseJsonSafely(raw)
  const nested = typeof parsed?.error === 'object' && parsed.error ? parsed.error : null
  const nestedObject = nested as Record<string, unknown> | null
  const topLevelError =
    typeof parsed?.error === 'string'
      ? parsed.error
      : typeof nestedObject?.error === 'string'
        ? nestedObject.error
        : null
  const description =
    typeof parsed?.error_description === 'string'
      ? parsed.error_description
      : typeof nestedObject?.message === 'string'
        ? nestedObject.message
        : null

  return new Error(description ?? topLevelError ?? raw ?? `Request failed with status ${response.status}.`)
}

const fetchWithRetry = async (input: RequestInfo | URL, init: RequestInit, retries = 2) => {
  let attempt = 0
  let lastError: Error | null = null

  while (attempt <= retries) {
    try {
      const response = await fetch(input, init)
      if (response.ok) return response

      if ([429, 500, 502, 503, 504].includes(response.status) && attempt < retries) {
        const retryAfterSeconds = Number(response.headers.get('retry-after') ?? '0')
        const retryDelay = Number.isFinite(retryAfterSeconds) && retryAfterSeconds > 0 ? retryAfterSeconds * 1000 : 1000 * (attempt + 1)
        await sleep(retryDelay)
        attempt += 1
        continue
      }

      throw await buildResponseError(response)
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error))
      if (attempt >= retries) {
        throw lastError
      }

      await sleep(1000 * (attempt + 1))
      attempt += 1
    }
  }

  throw lastError ?? new Error('Request failed.')
}

const fetchGraphCollection = async <T,>(url: string, accessToken: string) => {
  const items: T[] = []
  let nextUrl: string | undefined = url

  while (nextUrl) {
    const response = await fetchWithRetry(nextUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
      },
      method: 'GET',
    })
    const payload = (await response.json()) as GraphCollectionResponse<T>
    items.push(...(payload.value ?? []))
    nextUrl = payload['@odata.nextLink']
  }

  return items
}

const buildStableNotebookId = (sourceId: string) => `onenote-notebook:${sourceId}`
const buildStableGroupId = (sourceId: string) => `onenote-group:${sourceId}`
const buildStableSectionId = (sourceId: string) => `onenote-section:${sourceId}`
const buildStablePageId = (sourceId: string) => `onenote-page:${sourceId}`

const buildAssetId = (pageId: string, index: number) => `${ONE_NOTE_IMPORT_ASSET_PREFIX}${pageId}:${index}`

const buildFallbackFileName = (value: string, fallback: string) => {
  try {
    const url = new URL(value)
    const decoded = decodeURIComponent(url.pathname.split('/').filter(Boolean).at(-1) ?? '')
    return decoded || fallback
  } catch {
    return fallback
  }
}

const getMimeExtension = (mimeType: string) => {
  const normalized = mimeType.toLowerCase()
  if (normalized.includes('png')) return 'png'
  if (normalized.includes('jpeg') || normalized.includes('jpg')) return 'jpg'
  if (normalized.includes('gif')) return 'gif'
  if (normalized.includes('webp')) return 'webp'
  if (normalized.includes('bmp')) return 'bmp'
  if (normalized.includes('svg')) return 'svg'
  if (normalized.includes('pdf')) return 'pdf'
  if (normalized.includes('mpeg')) return 'mp3'
  if (normalized.includes('wav')) return 'wav'
  if (normalized.includes('ogg')) return 'ogg'
  if (normalized.includes('mp4')) return 'mp4'
  return 'bin'
}

const ensureFileExtension = (fileName: string, mimeType: string) => {
  if (/\.[a-z0-9]{2,5}$/i.test(fileName)) return fileName
  return `${fileName}.${getMimeExtension(mimeType)}`
}

const requestDeviceCode = async (clientId: string) => {
  const body = new URLSearchParams({
    client_id: clientId,
    scope: ONE_NOTE_SCOPES.join(' '),
  })

  const response = await fetchWithRetry(`${ONE_NOTE_AUTH_ROOT}/devicecode`, {
    body,
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    method: 'POST',
  })

  return (await response.json()) as DeviceCodeResponse
}

const pollForAccessToken = async (clientId: string, deviceCode: DeviceCodeResponse) => {
  const expiresAt = Date.now() + deviceCode.expires_in * 1000
  let pollDelay = Math.max(deviceCode.interval, 5) * 1000

  while (Date.now() < expiresAt) {
    await sleep(pollDelay)

    const body = new URLSearchParams({
      client_id: clientId,
      device_code: deviceCode.device_code,
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    })

    const response = await fetch(`${ONE_NOTE_AUTH_ROOT}/token`, {
      body,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      method: 'POST',
    })

    if (response.ok) {
      const payload = (await response.json()) as TokenResponse
      return payload.access_token
    }

    const payload = parseJsonSafely(await response.text())
    const errorCode = typeof payload?.error === 'string' ? payload.error : ''
    const errorDescription = typeof payload?.error_description === 'string' ? payload.error_description : ''

    if (errorCode === 'authorization_pending') {
      continue
    }

    if (errorCode === 'slow_down') {
      pollDelay += 5000
      continue
    }

    if (errorCode === 'authorization_declined') {
      throw new Error('Microsoft sign-in was cancelled before access was granted.')
    }

    if (errorCode === 'expired_token' || errorCode === 'bad_verification_code') {
      throw new Error('The Microsoft device code expired before sign-in completed.')
    }

    throw new Error(errorDescription || errorCode || 'Microsoft sign-in failed.')
  }

  throw new Error('Microsoft sign-in timed out before access was granted.')
}

const startMicrosoftDeviceSignIn = async (clientId: string) => {
  const deviceCode = await requestDeviceCode(clientId)
  const signInUrl = deviceCode.verification_uri_complete ?? deviceCode.verification_uri

  try {
    await navigator.clipboard.writeText(deviceCode.user_code)
  } catch {
    // Clipboard can fail in some runtime configurations; the sign-in code is still shown in the prompt.
  }

  window.open(signInUrl, '_blank', 'noopener,noreferrer')
  window.alert(
    [
      'A Microsoft sign-in page has been opened for OneNote import.',
      '',
      `Code: ${deviceCode.user_code}`,
      `URL: ${deviceCode.verification_uri}`,
      '',
      'The sign-in code has been copied if clipboard access is available.',
      'Finish the sign-in in your browser, then return to OnePlace and wait for the import to continue.',
    ].join('\n'),
  )

  return pollForAccessToken(clientId, deviceCode)
}

const fetchAssetFromUrl = async (
  assetUrl: string,
  accessToken: string,
  assetId: string,
  fallbackName: string,
  kind: AppAsset['kind'],
) => {
  const response = await fetchWithRetry(assetUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    method: 'GET',
  })
  const blob = await response.blob()
  const mimeType = blob.type || (kind === 'printout' ? 'application/pdf' : kind === 'image' ? 'image/png' : 'application/octet-stream')
  const dataUrl = await blobToDataUrl(blob)
  const name = ensureFileExtension(buildFallbackFileName(assetUrl, fallbackName), mimeType)

  return {
    asset: {
      createdAt: new Date().toISOString(),
      dataUrl,
      id: assetId,
      kind,
      mimeType,
      name,
      sizeLabel: formatSizeLabel(blob.size),
    } satisfies AppAsset,
    dataUrl,
  }
}

const createAttachmentCard = (doc: Document, asset: AppAsset) => {
  const card = doc.createElement('div')
  card.className = 'attachment-card'
  card.setAttribute('contenteditable', 'false')
  card.dataset.assetId = asset.id
  card.dataset.downloadUrl = asset.dataUrl
  card.dataset.fileName = asset.name

  const title = doc.createElement('strong')
  title.className = 'attachment-title'
  title.textContent = asset.name

  const meta = doc.createElement('span')
  meta.className = 'attachment-meta'
  meta.textContent = asset.sizeLabel

  card.append(title, meta)
  return card
}

const createPrintoutCard = (doc: Document, asset: AppAsset) => {
  const card = doc.createElement('section')
  card.className = 'printout-card'
  card.setAttribute('contenteditable', 'false')
  card.dataset.assetId = asset.id
  card.dataset.downloadUrl = asset.dataUrl
  card.dataset.fileName = asset.name

  const previewShell = doc.createElement('div')
  previewShell.className = 'printout-preview-shell'

  const frame = doc.createElement('iframe')
  frame.className = 'printout-preview'
  frame.dataset.assetId = asset.id
  frame.src = asset.dataUrl
  frame.title = asset.name

  const caption = doc.createElement('div')
  caption.className = 'printout-caption'
  caption.textContent = `${asset.name} · ${asset.sizeLabel}`

  previewShell.append(frame)
  card.append(previewShell, caption)
  return card
}

const createAudioNote = (doc: Document, asset: AppAsset) => {
  const figure = doc.createElement('figure')
  figure.className = 'audio-note'
  figure.setAttribute('contenteditable', 'false')

  const audio = doc.createElement('audio')
  audio.controls = true
  audio.dataset.assetId = asset.id
  audio.src = asset.dataUrl

  const caption = doc.createElement('figcaption')
  caption.textContent = asset.name

  figure.append(audio, caption)
  return figure
}

const replaceNodeWithTextFallback = (doc: Document, node: Element, text: string) => {
  const paragraph = doc.createElement('p')
  paragraph.textContent = text
  node.replaceWith(paragraph)
}

const importPageContent = async (pageId: string, accessToken: string) => {
  const response = await fetchWithRetry(`${ONE_NOTE_GRAPH_ROOT}/me/onenote/pages/${encodeURIComponent(pageId)}/content`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'text/html, application/xhtml+xml',
    },
    method: 'GET',
  })
  const rawHtml = await response.text()
  const baseUrl = response.url
  const doc = new DOMParser().parseFromString(rawHtml, 'text/html')
  const assets: AppAsset[] = []
  let assetIndex = 0

  const nextAssetId = () => buildAssetId(pageId, assetIndex++)
  const toAbsoluteUrl = (value: string) => {
    try {
      return new URL(value, baseUrl).toString()
    } catch {
      return value
    }
  }

  for (const image of [...doc.querySelectorAll('img')]) {
    const source = image.getAttribute('data-fullres-src') ?? image.getAttribute('src')
    if (!source) continue

    try {
      const { asset, dataUrl } = await fetchAssetFromUrl(
        toAbsoluteUrl(source),
        accessToken,
        nextAssetId(),
        image.getAttribute('alt')?.trim() || 'one-note-image',
        'image',
      )
      assets.push(asset)
      image.setAttribute('data-asset-id', asset.id)
      image.setAttribute('draggable', 'false')
      image.setAttribute('src', dataUrl)
    } catch {
      replaceNodeWithTextFallback(doc, image, image.getAttribute('alt')?.trim() || 'Image')
    }
  }

  for (const audio of [...doc.querySelectorAll('audio')]) {
    const source = audio.getAttribute('src')
    if (!source) continue

    try {
      const { asset } = await fetchAssetFromUrl(
        toAbsoluteUrl(source),
        accessToken,
        nextAssetId(),
        'one-note-audio',
        'audio',
      )
      assets.push(asset)
      audio.replaceWith(createAudioNote(doc, asset))
    } catch {
      replaceNodeWithTextFallback(doc, audio, 'Audio note')
    }
  }

  const embeddedFiles = [
    ...doc.querySelectorAll('object[data]'),
    ...doc.querySelectorAll('embed[src]'),
    ...doc.querySelectorAll('video[src]'),
    ...doc.querySelectorAll('iframe[src]'),
  ]

  for (const embedded of embeddedFiles) {
    const source =
      embedded.getAttribute('data') ??
      embedded.getAttribute('src') ??
      embedded.getAttribute('data-fullres-src')
    if (!source) continue

    const declaredType = embedded.getAttribute('type')?.toLowerCase() ?? ''
    const fallbackName =
      embedded.getAttribute('data-attachment')?.trim() ||
      embedded.getAttribute('title')?.trim() ||
      (declaredType.includes('pdf') ? 'one-note-printout.pdf' : 'one-note-attachment')

    try {
      const kind: AppAsset['kind'] = declaredType.includes('pdf') ? 'printout' : 'file'
      const { asset } = await fetchAssetFromUrl(toAbsoluteUrl(source), accessToken, nextAssetId(), fallbackName, kind)
      assets.push(asset)
      embedded.replaceWith(kind === 'printout' ? createPrintoutCard(doc, asset) : createAttachmentCard(doc, asset))
    } catch {
      replaceNodeWithTextFallback(doc, embedded, fallbackName)
    }
  }

  for (const node of [...doc.querySelectorAll('script, style, link, meta, title, base')]) {
    node.remove()
  }

  const sanitizedContent = sanitizePastedHtml(doc.body.innerHTML.trim() || '<p></p>')
  return {
    assets,
    content: sanitizedContent || '<p></p>',
  }
}

const mapWithConcurrency = async <T, R>(items: T[], limit: number, mapper: (item: T, index: number) => Promise<R>) => {
  if (items.length === 0) return [] as R[]

  const results = new Array<R>(items.length)
  let nextIndex = 0

  const worker = async () => {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex
      nextIndex += 1
      results[currentIndex] = await mapper(items[currentIndex], currentIndex)
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(limit, items.length) }, () => worker()),
  )

  return results
}

const nestPages = (pages: ImportedPageNode[]) => {
  const roots: Page[] = []
  const stack: Page[] = []

  for (const entry of pages) {
    const normalizedLevel = Math.max(0, Math.min(entry.level, stack.length))
    const node = { ...entry.page, children: [...entry.page.children] }

    while (stack.length > normalizedLevel) {
      stack.pop()
    }

    if (normalizedLevel === 0 || stack.length === 0) {
      roots.push(node)
    } else {
      stack[stack.length - 1].children.push(node)
    }

    stack[normalizedLevel] = node
    stack.length = normalizedLevel + 1
  }

  return roots
}

const createPlaceholderPage = (title: string, accent: string, pageId: string): Page => {
  const now = new Date().toISOString()
  const content = '<p></p>'
  return {
    accent,
    children: [],
    content,
    createdAt: now,
    id: pageId,
    inkStrokes: [],
    isCollapsed: false,
    snippet: buildSnippet(title, content, now),
    tags: [],
    task: null,
    title,
    updatedAt: now,
  }
}

const importSection = async (
  notebookName: string,
  sectionMeta: GraphSection,
  sectionColor: string,
  accessToken: string,
  onProgress: (message: string) => void,
) => {
  const pages = await fetchGraphCollection<GraphPage>(
    `${ONE_NOTE_GRAPH_ROOT}/me/onenote/sections/${encodeURIComponent(sectionMeta.id)}/pages?$top=100&pagelevel=true&select=id,title,createdDateTime,lastModifiedDateTime,order`,
    accessToken,
  )

  const orderedPages = [...pages].sort((left, right) => (left.order ?? 0) - (right.order ?? 0))
  const importedPages = await mapWithConcurrency(orderedPages, ONE_NOTE_PAGE_FETCH_CONCURRENCY, async (pageMeta, index) => {
    const pageTitle = pageMeta.title?.trim() || `Imported Page ${index + 1}`
    onProgress(`Importing ${notebookName} › ${sectionMeta.displayName} › ${pageTitle}`)
    const { assets, content } = await importPageContent(pageMeta.id, accessToken)
    const createdAt = pageMeta.createdDateTime ?? new Date().toISOString()
    const updatedAt = pageMeta.lastModifiedDateTime ?? createdAt
    const page: Page = {
      accent: sectionColor,
      children: [],
      content,
      createdAt,
      id: buildStablePageId(pageMeta.id),
      inkStrokes: [],
      isCollapsed: false,
      snippet: buildSnippet(pageTitle, stripHtml(content) ? content : '<p></p>', updatedAt),
      tags: [],
      task: null,
      title: pageTitle,
      updatedAt,
    }

    return {
      assets,
      level: pageMeta.level ?? 0,
      page,
    } satisfies ImportedPageNode
  })

  const nestedPages = nestPages(importedPages)
  const assets = importedPages.flatMap((entry) => entry.assets)
  const sectionPages =
    nestedPages.length > 0
      ? nestedPages
      : [createPlaceholderPage(`Imported notes for ${sectionMeta.displayName}`, sectionColor, `${buildStableSectionId(sectionMeta.id)}:placeholder`)]

  return {
    assets,
    section: {
      color: sectionColor,
      id: buildStableSectionId(sectionMeta.id),
      name: sectionMeta.displayName || 'Imported Section',
      pages: sectionPages,
      passwordHash: null,
      passwordHint: '',
    } satisfies Section,
  }
}

const importSections = async (
  notebookName: string,
  sectionMetas: GraphSection[],
  sectionGroupName: string,
  accessToken: string,
  colorOffset: number,
  onProgress: (message: string) => void,
) => {
  const sections: Section[] = []
  const assets: AppAsset[] = []

  for (let index = 0; index < sectionMetas.length; index += 1) {
    const color = accentPalette[(colorOffset + index) % accentPalette.length]
    const imported = await importSection(notebookName, sectionMetas[index], color, accessToken, onProgress)
    sections.push(imported.section)
    assets.push(...imported.assets)
  }

  const safeSections =
    sections.length > 0
      ? sections
      : [
          {
            color: accentPalette[colorOffset % accentPalette.length],
            id: `${buildStableGroupId(sectionGroupName)}:section`,
            name: sectionGroupName,
            pages: [createPlaceholderPage(`Imported notes for ${sectionGroupName}`, accentPalette[colorOffset % accentPalette.length], `${buildStableGroupId(sectionGroupName)}:page`)],
            passwordHash: null,
            passwordHint: '',
          } satisfies Section,
        ]

  return {
    assets,
    sectionGroup: {
      id: buildStableGroupId(sectionGroupName),
      isCollapsed: false,
      name: sectionGroupName,
      sections: safeSections,
    } satisfies SectionGroup,
  }
}

const importSectionGroupsRecursively = async (
  notebookName: string,
  notebookId: string,
  groupMeta: GraphSectionGroup,
  pathNames: string[],
  accessToken: string,
  colorOffset: number,
  onProgress: (message: string) => void,
) => {
  const importedGroups: SectionGroup[] = []
  const assets: AppAsset[] = []
  const directSections = await fetchGraphCollection<GraphSection>(
    `${ONE_NOTE_GRAPH_ROOT}/me/onenote/sectionGroups/${encodeURIComponent(groupMeta.id)}/sections?$top=100&select=id,displayName`,
    accessToken,
  )

  if (directSections.length > 0) {
    const imported = await importSections(
      notebookName,
      directSections,
      pathNames.join(' / '),
      accessToken,
      colorOffset,
      onProgress,
    )
    importedGroups.push({
      ...imported.sectionGroup,
      id: buildStableGroupId(`${notebookId}:${groupMeta.id}`),
    })
    assets.push(...imported.assets)
  }

  const childGroups = await fetchGraphCollection<GraphSectionGroup>(
    `${ONE_NOTE_GRAPH_ROOT}/me/onenote/sectionGroups/${encodeURIComponent(groupMeta.id)}/sectionGroups?$top=100&select=id,displayName`,
    accessToken,
  )

  for (let index = 0; index < childGroups.length; index += 1) {
    const nested = await importSectionGroupsRecursively(
      notebookName,
      notebookId,
      childGroups[index],
      [...pathNames, childGroups[index].displayName || `Group ${index + 1}`],
      accessToken,
      colorOffset + index + 1,
      onProgress,
    )
    importedGroups.push(...nested.sectionGroups)
    assets.push(...nested.assets)
  }

  return {
    assets,
    sectionGroups: importedGroups,
  }
}

const importNotebook = async (
  notebookMeta: GraphNotebook,
  notebookIndex: number,
  accessToken: string,
  onProgress: (message: string) => void,
) => {
  const notebookName = notebookMeta.displayName?.trim() || `Imported Notebook ${notebookIndex + 1}`
  const assets: AppAsset[] = []
  const sectionGroups: SectionGroup[] = []

  const directSections = await fetchGraphCollection<GraphSection>(
    `${ONE_NOTE_GRAPH_ROOT}/me/onenote/notebooks/${encodeURIComponent(notebookMeta.id)}/sections?$top=100&select=id,displayName`,
    accessToken,
  )

  if (directSections.length > 0) {
    const imported = await importSections(
      notebookName,
      directSections,
      'Sections',
      accessToken,
      notebookIndex,
      onProgress,
    )
    sectionGroups.push({
      ...imported.sectionGroup,
      id: buildStableGroupId(`${notebookMeta.id}:root`),
    })
    assets.push(...imported.assets)
  }

  const rootGroups = await fetchGraphCollection<GraphSectionGroup>(
    `${ONE_NOTE_GRAPH_ROOT}/me/onenote/notebooks/${encodeURIComponent(notebookMeta.id)}/sectionGroups?$top=100&select=id,displayName`,
    accessToken,
  )

  for (let groupIndex = 0; groupIndex < rootGroups.length; groupIndex += 1) {
    const nested = await importSectionGroupsRecursively(
      notebookName,
      notebookMeta.id,
      rootGroups[groupIndex],
      [rootGroups[groupIndex].displayName || `Group ${groupIndex + 1}`],
      accessToken,
      notebookIndex + groupIndex + 1,
      onProgress,
    )
    sectionGroups.push(...nested.sectionGroups)
    assets.push(...nested.assets)
  }

  if (sectionGroups.length === 0) {
    const color = accentPalette[notebookIndex % accentPalette.length]
    sectionGroups.push(
      createSectionGroup('Sections', [
        {
          color,
          id: `${buildStableNotebookId(notebookMeta.id)}:section`,
          name: 'Imported Section',
          pages: [createPlaceholderPage('Imported Page', color, `${buildStableNotebookId(notebookMeta.id)}:page`)],
          passwordHash: null,
          passwordHint: '',
        },
      ]),
    )
  }

  return {
    assets,
    notebook: {
      color: accentPalette[notebookIndex % accentPalette.length],
      icon: 'book',
      id: buildStableNotebookId(notebookMeta.id),
      name: notebookName,
      sectionGroups,
    } satisfies Notebook,
  }
}

const importAllAccessibleNotebooks = async (accessToken: string, onProgress: (message: string) => void) => {
  const notebooks = await fetchGraphCollection<GraphNotebook>(
    `${ONE_NOTE_GRAPH_ROOT}/me/onenote/notebooks?$top=100&select=id,displayName`,
    accessToken,
  )

  const importedNotebooks: Notebook[] = []
  const importedAssets: AppAsset[] = []

  for (let index = 0; index < notebooks.length; index += 1) {
    onProgress(`Reading notebook ${index + 1} of ${notebooks.length}: ${notebooks[index].displayName}`)
    const imported = await importNotebook(notebooks[index], index, accessToken, onProgress)
    importedNotebooks.push(imported.notebook)
    importedAssets.push(...imported.assets)
  }

  return {
    assets: importedAssets,
    notebooks: importedNotebooks,
  }
}

const mergeImportedNotebooksIntoState = (current: AppState, importedNotebooks: Notebook[], importedAssets: AppAsset[]) => {
  const importedNotebookIds = new Set(importedNotebooks.map((item) => item.id))
  const nextNotebooks = [...current.notebooks.filter((item) => !importedNotebookIds.has(item.id)), ...importedNotebooks]
  const preservedAssets = Object.fromEntries(
    Object.entries(current.meta.assets).filter(([assetId]) => !assetId.startsWith(ONE_NOTE_IMPORT_ASSET_PREFIX)),
  )
  const nextAssets = {
    ...preservedAssets,
    ...Object.fromEntries(importedAssets.map((asset) => [asset.id, asset])),
  }
  const firstImportedNotebook = importedNotebooks[0]
  const firstImportedGroup = firstImportedNotebook?.sectionGroups[0]
  const firstImportedSection = firstImportedGroup?.sections[0]
  const firstImportedPage = firstImportedSection?.pages[0]

  return ensureSelection({
    ...current,
    meta: {
      ...current.meta,
      assets: nextAssets,
    },
    notebooks: nextNotebooks,
    selectedNotebookId: firstImportedNotebook?.id ?? current.selectedNotebookId,
    selectedPageId: firstImportedPage?.id ?? current.selectedPageId,
    selectedSectionGroupId: firstImportedGroup?.id ?? current.selectedSectionGroupId,
    selectedSectionId: firstImportedSection?.id ?? current.selectedSectionId,
  })
}

type UseOneNoteImportArgs = {
  setActiveTab: (tab: 'Home') => void
  setAppState: Dispatch<SetStateAction<AppState>>
  setSaveLabel: (value: string) => void
}

export const useOneNoteImport = ({ setActiveTab, setAppState, setSaveLabel }: UseOneNoteImportArgs) => {
  const [isImportingOneNote, setIsImportingOneNote] = useState(false)

  const importAllOneNoteNotebooks = async () => {
    if (isImportingOneNote) return

    const clientId = ONEPLACE_ONENOTE_CLIENT_ID.trim()
    if (!clientId) {
      const message =
        'OneNote import is not configured yet. Add the official OnePlace Microsoft Client ID in src/app/appConfig.ts before shipping this feature.'
      setSaveLabel(message)
      window.alert(message)
      return
    }

    const shouldImport = window.confirm(
      'Import every notebook available in this Microsoft OneNote account into OnePlace now?\n\nExisting OneNote-imported notebooks in OnePlace will be refreshed by ID on re-import.',
    )
    if (!shouldImport) return

    setIsImportingOneNote(true)
    try {
      setSaveLabel('Waiting for Microsoft sign-in...')
      const accessToken = await startMicrosoftDeviceSignIn(clientId)
      setSaveLabel('Connected to Microsoft. Reading OneNote notebooks...')

      const imported = await importAllAccessibleNotebooks(accessToken, (message) => {
        setSaveLabel(message)
      })

      if (imported.notebooks.length === 0) {
        setSaveLabel('No OneNote notebooks were found for this account')
        window.alert('Microsoft sign-in worked, but no OneNote notebooks were returned for this account.')
        return
      }

      setAppState((current) => mergeImportedNotebooksIntoState(current, imported.notebooks, imported.assets))
      setActiveTab('Home')
      setSaveLabel(
        `Imported ${imported.notebooks.length} OneNote notebook${imported.notebooks.length === 1 ? '' : 's'} into OnePlace`,
      )
    } catch (error) {
      const message = normalizeOneNoteImportError(error)
      setSaveLabel(`OneNote import failed: ${message}`)
      window.alert(`OneNote import failed.\n\n${message}`)
    } finally {
      setIsImportingOneNote(false)
    }
  }

  return {
    importAllOneNoteNotebooks,
    isImportingOneNote,
  }
}
