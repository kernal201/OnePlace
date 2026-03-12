import { invoke } from '@tauri-apps/api/core'
import { open } from '@tauri-apps/plugin-dialog'

const DATA_KEY = 'oneplace-data'

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
