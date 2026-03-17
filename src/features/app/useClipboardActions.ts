import { sanitizePastedHtml } from '../../app/appModel'

type Args = {
  handleImageSelection: (files: File[]) => Promise<void>
  insertHtmlAtSelection: (html: string) => void
  insertTextAsHtml: (value: string) => void
  setSaveLabel: (value: string) => void
}

export const useClipboardActions = ({
  handleImageSelection,
  insertHtmlAtSelection,
  insertTextAsHtml,
  setSaveLabel,
}: Args) => {
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

  return { pasteFromClipboard }
}
