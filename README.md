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

End users do not need this repo, Node.js, Rust, or the `src-tauri/...` folder.

They only need the installer `.exe`.

Users should download the installer from the GitHub Releases page for this repo:

- `https://github.com/ShadowKernal/OnePlace/releases`

Ship the generated NSIS installer to users. They can install it by:

1. Downloading `OnePlace_<version>_x64-setup.exe`
2. Saving it anywhere on their PC, such as `Downloads`
3. Double-clicking the installer
3. Launching `OnePlace` from the Start menu

The installer is configured for per-user installation, so it should not require admin access in the normal case.

## What To Send People

Send this installer file to Windows users:

- `src-tauri/target/release/bundle/nsis/OnePlace_0.1.0_x64-setup.exe`

Do not send the unpacked `oneplace.exe` from `src-tauri/target/release/`. The correct file to distribute is the `-setup.exe` installer.

Recommended distribution options:

- Upload the installer to a GitHub Release
- Send the installer file directly
- Upload it to Google Drive, Dropbox, or another file host

If someone downloads `OnePlace_0.1.0_x64-setup.exe`, they can run it directly on Windows.

Recommended public download location:

- `https://github.com/ShadowKernal/OnePlace/releases`

## macOS Release Notes

macOS builds are architecture-specific.

- Intel Macs need the `x64` macOS release asset
- Apple Silicon Macs need the `aarch64` macOS release asset

If you publish only one macOS `.dmg`, one class of Mac users will get a build that does not launch.

## Notes

- Tauri bundling is enabled in `src-tauri/tauri.conf.json`.
- The current installer target is `nsis`.
- If you later want MSI output too, add it to `bundle.targets` and rebuild.
