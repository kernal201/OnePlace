import type { Dispatch, SetStateAction } from 'react'
import {
  buildSnippet,
  defaultTask,
  updateNestedPages,
} from '../../app/appModel'
import type { AppState, Page } from '../../app/appModel'

export const updatePageById = (
  setAppState: Dispatch<SetStateAction<AppState>>,
  pageId: string,
  updater: (targetPage: Page) => Page,
) => {
  setAppState((current) => ({
    ...current,
    notebooks: current.notebooks.map((item) => ({
      ...item,
      sectionGroups: item.sectionGroups.map((group) => ({
        ...group,
        sections: group.sections.map((entry) => ({
          ...entry,
          pages: updateNestedPages(entry.pages, pageId, updater),
        })),
      })),
    })),
  }))
}

export const buildTaskToggleUpdate = (note: Page) => ({
  ...note,
  snippet: buildSnippet(note.title, note.content),
  task: {
    ...note.task!,
    status: note.task?.status === 'done' ? 'open' : 'done',
  },
  updatedAt: new Date().toISOString(),
})

export const buildTaskDueUpdate = (note: Page, value: string) => ({
  ...note,
  snippet: buildSnippet(note.title, note.content),
  task: note.task
    ? { ...note.task, dueAt: value ? new Date(value).toISOString() : null }
    : { ...defaultTask(), dueAt: value ? new Date(value).toISOString() : null },
  updatedAt: new Date().toISOString(),
})
