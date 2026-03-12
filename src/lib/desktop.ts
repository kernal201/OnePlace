import { invoke } from '@tauri-apps/api/core'
import { open } from '@tauri-apps/plugin-dialog'
import { relaunch } from '@tauri-apps/plugin-process'
import { check } from '@tauri-apps/plugin-updater'

const DATA_KEY = 'oneplace-data'

export type DesktopUpdateInfo = {
  body?: string
  currentVersion: string
  version: string
}

export type DesktopUpdateProgress =
  | { event: 'Started'; data: { contentLength: number | null } }
  | { event: 'Progress'; data: { chunkLength: number } }
  | { event: 'Finished' }

const isTauriRuntime = () =>
  typeof window !== 'undefined' &&
  (Object.prototype.hasOwnProperty.call(window, '__TAURI_INTERNALS__') ||
    Object.prototype.hasOwnProperty.call(window, '__TAURI__'))

export const loadDesktopData = async (): Promise<string | null> => {
  if (isTauriRuntime()) {
    try {
      return await invoke<string | null>('load_data')
    } catch {
      return localStorage.getItem(DATA_KEY)
    }
  }

  return localStorage.getItem(DATA_KEY)
}

export const saveDesktopData = async (rawData: string): Promise<DesktopSaveResult> => {
  if (isTauriRuntime()) {
    try {
      return await invoke<DesktopSaveResult>('save_data', { rawData })
    } catch {
      localStorage.setItem(DATA_KEY, rawData)
      return {
        path: 'localStorage',
        savedAt: new Date().toISOString(),
      }
    }
  }

  localStorage.setItem(DATA_KEY, rawData)
  return {
    path: 'localStorage',
    savedAt: new Date().toISOString(),
  }
}

export const getDesktopAppInfo = async (): Promise<DesktopAppInfo | null> => {
  if (isTauriRuntime()) {
    try {
      return await invoke<DesktopAppInfo>('get_app_info')
    } catch {
      return null
    }
  }

  return null
}

export const checkForDesktopUpdate = async (): Promise<DesktopUpdateInfo | null> => {
  if (!isTauriRuntime()) return null

  const update = await check()
  if (!update) return null

  return {
    body: update.body ?? undefined,
    currentVersion: update.currentVersion,
    version: update.version,
  }
}

export const downloadAndInstallDesktopUpdate = async (
  onEvent?: (event: DesktopUpdateProgress) => void,
): Promise<void> => {
  if (!isTauriRuntime()) return

  const update = await check()
  if (!update) return

  await update.downloadAndInstall((event) => {
    if (event.event === 'Started') {
      onEvent?.({
        event: 'Started',
        data: {
          contentLength: event.data.contentLength ?? null,
        },
      })
      return
    }

    if (event.event === 'Progress') {
      onEvent?.({
        event: 'Progress',
        data: {
          chunkLength: event.data.chunkLength,
        },
      })
      return
    }

    onEvent?.({ event: 'Finished' })
  })

  await relaunch()
}

export const pickNotebookDirectory = async (): Promise<string | null> => {
  if (!isTauriRuntime()) return null

  const selected = await open({
    directory: true,
    multiple: false,
    title: 'Open Notebook Folder',
  })

  return typeof selected === 'string' ? selected : null
}

export const openNotebookDirectory = async (path: string): Promise<string> =>
  invoke<string>('open_notebook_dir', { path })

export const exportNotebookDirectory = async (
  path: string,
  notebook: string,
): Promise<DesktopSaveResult> => invoke<DesktopSaveResult>('export_notebook_dir', { notebook, path })
