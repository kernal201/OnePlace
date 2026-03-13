# Changelog

## 0.1.6 - 2026-03-13

- Fixed drag-and-drop for visible sections in the left notebook pane.
- Added drop indicators for section reordering.

## 0.1.5 - 2026-03-12

- Moved the visible app version into the note header beside the page action buttons.
- Kept the status-strip version fallback, but no longer rely on the footer for discoverability.

## 0.1.4 - 2026-03-12

- Always show the app version in the bottom-right status strip.
- Added a frontend build-time version fallback when desktop app info is unavailable.

## 0.1.3 - 2026-03-12

- Fixed automatic update checks so they run even when desktop app metadata is unavailable.
- Added a manual `Check for updates` action in the left pane footer.
- Improved update status messages for no-update and failure cases.

## 0.1.2 - 2026-03-12

- Added in-app update prompts that show the latest release notes before install.
- Published Intel and Apple Silicon macOS release artifacts side by side.
- Added signed updater release automation for Windows and macOS builds.

## 0.1.1 - 2026-03-12

- Added updater support to the Tauri desktop shell.
- Added GitHub Releases publishing for signed updater artifacts.
- Added architecture-specific macOS release packaging.
