interface DesktopSaveResult {
  path: string;
  savedAt: string;
}

interface DesktopAppInfo {
  dataPath: string;
  name: string;
  version: string;
}

interface ImportedOneNoteFile {
  absolutePath: string;
  contents: string;
  extension: string;
  modifiedAt: string;
  name: string;
  relativeDir: string;
  relativePath: string;
}

interface ImportedOneNoteDirectory {
  files: ImportedOneNoteFile[];
  name: string;
  path: string;
}

interface ImportedAssetData {
  dataUrl: string;
  mimeType: string;
  name: string;
  size: number;
}

interface SpeechRecognitionAlternative {
  transcript: string;
  confidence: number;
}

interface SpeechRecognitionResult {
  isFinal: boolean;
  length: number;
  [index: number]: SpeechRecognitionAlternative;
}

interface SpeechRecognitionResultList {
  length: number;
  [index: number]: SpeechRecognitionResult;
}

interface SpeechRecognitionEvent extends Event {
  resultIndex: number;
  results: SpeechRecognitionResultList;
}

interface SpeechRecognitionErrorEvent extends Event {
  error: string;
  message: string;
}

interface BrowserSpeechRecognition extends EventTarget {
  continuous: boolean;
  interimResults: boolean;
  lang: string;
  onend: ((this: BrowserSpeechRecognition, ev: Event) => unknown) | null;
  onerror: ((this: BrowserSpeechRecognition, ev: SpeechRecognitionErrorEvent) => unknown) | null;
  onresult: ((this: BrowserSpeechRecognition, ev: SpeechRecognitionEvent) => unknown) | null;
  start(): void;
  stop(): void;
}

interface BrowserSpeechRecognitionConstructor {
  new (): BrowserSpeechRecognition;
}

interface Window {
  SpeechRecognition?: BrowserSpeechRecognitionConstructor;
  webkitSpeechRecognition?: BrowserSpeechRecognitionConstructor;
}
