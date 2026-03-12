# OnePlace

OnePlace is a desktop note-taking app built with React, Vite, TypeScript, and Tauri.

## Development

Requirements:

- Node.js 20+
- Rust toolchain
- Visual Studio C++ build tools on Windows

Run locally:

```bash
npm install
npm run dev
```

The Tauri shell will start the desktop app and use the Vite frontend in dev mode.

## Build A Windows Installer

This project is configured to generate a Windows NSIS installer.

Build it with:

```bash
npm install
npm run build:installer
```

Installer output:

- `src-tauri/target/release/bundle/nsis/OnePlace_<version>_x64-setup.exe`

The unpacked executable is also produced under:

- `src-tauri/target/release/oneplace.exe`

## Install For End Users

Ship the generated NSIS installer to users. They can install it by:

1. Downloading `OnePlace_<version>_x64-setup.exe`
2. Running the installer
3. Launching `OnePlace` from the Start menu

The installer is configured for per-user installation, so it should not require admin access in the normal case.

## Notes

- Tauri bundling is enabled in `src-tauri/tauri.conf.json`.
- The current installer target is `nsis`.
- If you later want MSI output too, add it to `bundle.targets` and rebuild.
