import type { ReactNode } from 'react'

type IconProps = {
  className?: string
  size?: number
}

const Svg = ({
  children,
  className,
  size = 18,
  viewBox = '0 0 24 24',
}: IconProps & { children: ReactNode; viewBox?: string }) => (
  <svg
    aria-hidden="true"
    className={className}
    fill="none"
    height={size}
    viewBox={viewBox}
    width={size}
    xmlns="http://www.w3.org/2000/svg"
  >
    {children}
  </svg>
)

export const OneNoteLogoIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <rect fill="#7719aa" height="20" rx="3" width="20" x="2" y="2" />
    <path d="M8 6.5H16V17.5H8V6.5Z" fill="#8d3dbe" />
    <path d="M6.5 7.5H13.5V16.5H6.5V7.5Z" fill="white" />
    <path d="M8.5 14.8V9.2L11.8 14.8V9.2" stroke="#7719aa" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.4" />
  </Svg>
)

export const SaveIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M5 4.5H16.8L19 6.8V19.5H5V4.5Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M8 4.5H15V9H8V4.5Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M8 19V13H16V19" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
  </Svg>
)

export const UndoIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M9 7H5V3" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.7" />
    <path d="M5.5 7.2C7.2 5.2 9.4 4.2 12.2 4.2C16.3 4.2 19.4 7.1 19.4 11C19.4 14.9 16.3 17.8 12.2 17.8C9.8 17.8 7.8 16.9 6.2 15.4" stroke="currentColor" strokeLinecap="round" strokeWidth="1.7" />
  </Svg>
)

export const ChevronDownIcon = ({ className, size = 14 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M6 9L12 15L18 9" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.7" />
  </Svg>
)

export const SearchIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <circle cx="11" cy="11" r="5.5" stroke="currentColor" strokeWidth="1.7" />
    <path d="M15.2 15.2L19 19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.7" />
  </Svg>
)

export const NotebookStackIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <rect height="17" rx="2.4" stroke="currentColor" strokeWidth="1.5" width="13" x="6.5" y="3.5" />
    <path d="M9 6.5V17.5" stroke="currentColor" strokeWidth="1.5" />
    <path d="M3.5 6.5V17.5" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const FolderIcon = ({ className, color = '#d9aa35', size = 18 }: IconProps & { color?: string }) => (
  <Svg className={className} size={size}>
    <path d="M3.5 7.5H9L10.8 5.5H20.5V17.8C20.5 18.7 19.7 19.5 18.8 19.5H5.2C4.3 19.5 3.5 18.7 3.5 17.8V7.5Z" fill={color} stroke="rgba(0,0,0,0.18)" strokeWidth="1" />
    <path d="M3.5 8.3H20.5" stroke="rgba(255,255,255,0.55)" strokeWidth="1" />
  </Svg>
)

export const PersonIcon = ({ className, color = '#2f9b90', size = 18 }: IconProps & { color?: string }) => (
  <Svg className={className} size={size}>
    <circle cx="12" cy="8.2" fill={color} r="3.2" />
    <path d="M5 18C5.9 14.8 8.5 13 12 13C15.5 13 18.1 14.8 19 18" fill={color} stroke={color} strokeLinecap="round" strokeWidth="1.2" />
  </Svg>
)

export const SectionBookIcon = ({ className, color = '#7e42b3', size = 18 }: IconProps & { color?: string }) => (
  <Svg className={className} size={size}>
    <rect fill={color} height="18" rx="2.2" width="15" x="4.5" y="3" />
    <rect fill="rgba(255,255,255,0.2)" height="18" rx="1.6" width="3" x="7" y="3" />
    <rect fill="white" height="8" opacity="0.85" rx="1" width="6" x="11" y="8" />
  </Svg>
)

export const ListLinesIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M6 7H19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.7" />
    <path d="M6 12H19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.7" />
    <path d="M6 17H19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.7" />
    <circle cx="3.5" cy="7" fill="currentColor" r="1" />
    <circle cx="3.5" cy="12" fill="currentColor" r="1" />
    <circle cx="3.5" cy="17" fill="currentColor" r="1" />
  </Svg>
)

export const EditIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M5 16.8L8.2 16.2L18 6.5L15.5 4L5.8 13.8L5 16.8Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M14.6 4.8L17.1 7.3" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const SettingsIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <circle cx="12" cy="12" r="3.1" stroke="currentColor" strokeWidth="1.6" />
    <path d="M12 4V6.2M12 17.8V20M4 12H6.2M17.8 12H20M6.3 6.3L7.8 7.8M16.2 16.2L17.7 17.7M17.7 6.3L16.2 7.8M7.8 16.2L6.3 17.7" stroke="currentColor" strokeLinecap="round" strokeWidth="1.6" />
  </Svg>
)

export const PasteIcon = ({ className, size = 28 }: IconProps) => (
  <Svg className={className} size={size}>
    <rect height="18" rx="2.5" stroke="currentColor" strokeWidth="1.5" width="14" x="6" y="5" />
    <path d="M9 5.2V3.8C9 3.4 9.4 3 9.8 3H14.2C14.6 3 15 3.4 15 3.8V5.2" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M4 9.5H8.5" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const CutIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <circle cx="5.5" cy="17" r="2.3" stroke="currentColor" strokeWidth="1.5" />
    <circle cx="5.5" cy="7" r="2.3" stroke="currentColor" strokeWidth="1.5" />
    <path d="M8 8.6L18 17.2" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M8 15.4L18 6.8" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const CopyIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <rect height="10" rx="1.5" stroke="currentColor" strokeWidth="1.5" width="8" x="9" y="4" />
    <rect height="10" rx="1.5" stroke="currentColor" strokeWidth="1.5" width="8" x="6" y="9" />
  </Svg>
)

export const BrushIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M12 5L18 11L12.8 16.2L6.8 10.2L12 5Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M6.8 10.2L4 13C4 15.2 5.2 16.4 7.4 16.4L10.2 13.6" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
  </Svg>
)

export const TextSizeUpIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M6 18L10.2 6H10.8L15 18M7.4 14H13.6" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M18.2 7V11.4M16 9.2H20.4" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const TextSizeDownIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M6 18L10.2 6H10.8L15 18M7.4 14H13.6" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M16.2 9.2H20.2" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const PenIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M5 19L8.8 18L18.3 8.5L14.5 4.7L5 14.2L5 19Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M13.8 5.4L17.6 9.2" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const BoldIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M8 5H13.2C15.2 5 16.5 6.1 16.5 7.9C16.5 9.5 15.4 10.5 13.8 10.7C15.8 10.9 17 12.1 17 14C17 16 15.5 17.2 13.1 17.2H8V5Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.6" />
  </Svg>
)

export const ItalicIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M13.5 5H10.5M13.5 19H10.5M13 5L11 19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.6" />
  </Svg>
)

export const UnderlineIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M8 5V11C8 13.4 9.6 15 12 15C14.4 15 16 13.4 16 11V5" stroke="currentColor" strokeLinecap="round" strokeWidth="1.6" />
    <path d="M7 19H17" stroke="currentColor" strokeLinecap="round" strokeWidth="1.6" />
  </Svg>
)

export const BulletsIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <circle cx="4" cy="7" fill="currentColor" r="1.2" />
    <circle cx="4" cy="12" fill="currentColor" r="1.2" />
    <circle cx="4" cy="17" fill="currentColor" r="1.2" />
    <path d="M8 7H19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M8 12H19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M8 17H19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const IndentIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M9 7H20M9 12H17M9 17H20" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M4 12H8M6 10L4 12L6 14" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
  </Svg>
)

export const SubpageIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M4 7H15M4 12H13M4 17H15" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M16 12H20M18 10L20 12L18 14" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
  </Svg>
)

export const AlignLeftIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M5 6H19M5 10H15M5 14H19M5 18H13" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const SortIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M8 6V18M8 18L5.5 15.5M8 18L10.5 15.5" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M14 6H18M14 11H19M14 16H17" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const ShowIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M3.5 12C5.5 8.5 8.4 6.8 12 6.8C15.6 6.8 18.5 8.5 20.5 12C18.5 15.5 15.6 17.2 12 17.2C8.4 17.2 5.5 15.5 3.5 12Z" stroke="currentColor" strokeWidth="1.5" />
    <circle cx="12" cy="12" r="2.3" stroke="currentColor" strokeWidth="1.5" />
  </Svg>
)

export const LinkIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M9.2 14.8L7.1 16.9C5.6 18.4 3.4 18.4 1.9 16.9C0.4 15.4 0.4 13.2 1.9 11.7L5.4 8.2C6.9 6.7 9.1 6.7 10.6 8.2" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M14.8 9.2L16.9 7.1C18.4 5.6 20.6 5.6 22.1 7.1C23.6 8.6 23.6 10.8 22.1 12.3L18.6 15.8C17.1 17.3 14.9 17.3 13.4 15.8" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M8.5 15.5L15.5 8.5" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)

export const ImageIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <rect height="15" rx="1.8" stroke="currentColor" strokeWidth="1.5" width="18" x="3" y="4.5" />
    <circle cx="8" cy="9.5" fill="currentColor" r="1.3" />
    <path d="M5.5 16L10.5 11L14 14.5L16.3 12.2L18.5 14.4" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
  </Svg>
)

export const AttachmentIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M8 12.5L14.8 5.7C16 4.5 17.9 4.5 19.1 5.7C20.3 6.9 20.3 8.8 19.1 10L10.4 18.7C8.8 20.3 6.3 20.3 4.7 18.7C3.1 17.1 3.1 14.6 4.7 13L12.8 4.9" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
  </Svg>
)

export const TableIcon = ({ className, size = 18 }: IconProps) => (
  <Svg className={className} size={size}>
    <rect height="15" rx="1.4" stroke="currentColor" strokeWidth="1.5" width="18" x="3" y="4.5" />
    <path d="M3.8 9.5H20.2M3.8 14.5H20.2M9 5.3V18.7M15 5.3V18.7" stroke="currentColor" strokeWidth="1.4" />
  </Svg>
)

export const InsertFormattingIcon = ({ className, size = 26 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M4.5 19.5L10.5 4.5H13.5L19.5 19.5M7 14H17" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M15.5 17.2L20 12.8" stroke="#61a7d8" strokeLinecap="round" strokeWidth="2" />
  </Svg>
)

export const FormatMotivationIcon = ({ className, size = 26 }: IconProps) => (
  <Svg className={className} size={size}>
    <rect height="14" rx="1.7" stroke="currentColor" strokeWidth="1.5" width="12" x="5" y="5" />
    <path d="M17 9H21" stroke="#61a7d8" strokeLinecap="round" strokeWidth="2" />
    <path d="M17 13H21" stroke="#61a7d8" strokeLinecap="round" strokeWidth="2" />
    <path d="M9 5V19" stroke="currentColor" strokeWidth="1.5" />
    <path d="M5 10H17M5 14H17" stroke="currentColor" strokeWidth="1.5" />
  </Svg>
)

export const ProjectIcon = ({ className, size = 26 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M6 6.5H11L12.5 8H18V18.5H6V6.5Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M9 11H15.5" stroke="#61a7d8" strokeLinecap="round" strokeWidth="1.8" />
    <path d="M9 14H13.5" stroke="#61a7d8" strokeLinecap="round" strokeWidth="1.8" />
  </Svg>
)

export const TagsIcon = ({ className, size = 26 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M6 8.5L12.5 2H18L22 6V11.5L15.5 18L6 8.5Z" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <circle cx="16.8" cy="7.2" fill="currentColor" r="1.2" />
    <path d="M10 12.5L14.2 8.3" stroke="#61a7d8" strokeLinecap="round" strokeWidth="1.8" />
  </Svg>
)

export const DeleteIcon = ({ className, size = 16 }: IconProps) => (
  <Svg className={className} size={size}>
    <path d="M5 7H19" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M9 7V18M15 7V18" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
    <path d="M7.5 7L8.2 19.2H15.8L16.5 7" stroke="currentColor" strokeLinejoin="round" strokeWidth="1.5" />
    <path d="M9.2 4.5H14.8" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
  </Svg>
)
