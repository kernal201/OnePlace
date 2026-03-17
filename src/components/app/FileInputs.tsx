import type { ChangeEvent as ReactChangeEvent, RefObject } from 'react'

type FileInputsProps = {
  imageInputRef: RefObject<HTMLInputElement | null>
  attachmentInputRef: RefObject<HTMLInputElement | null>
  printoutInputRef: RefObject<HTMLInputElement | null>
  handleImageSelection: (files: File[]) => void | Promise<void>
  handleAttachmentSelection: (files: File[]) => void | Promise<void>
  handlePrintoutSelection: (files: File[]) => void | Promise<void>
}

export function FileInputs({
  imageInputRef,
  attachmentInputRef,
  printoutInputRef,
  handleImageSelection,
  handleAttachmentSelection,
  handlePrintoutSelection,
}: FileInputsProps) {
  return (
    <>
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
    </>
  )
}
