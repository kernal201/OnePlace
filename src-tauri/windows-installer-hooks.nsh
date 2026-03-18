!macro RefreshShortcutIfPresent SHORTCUT_PATH
  IfFileExists "${SHORTCUT_PATH}" 0 +4
    Delete "${SHORTCUT_PATH}"
    CreateShortcut "${SHORTCUT_PATH}" "$INSTDIR\${MAINBINARYNAME}.exe"
    !insertmacro SetLnkAppUserModelId "${SHORTCUT_PATH}"
!macroend

!macro NSIS_HOOK_POSTINSTALL
  !if "${STARTMENUFOLDER}" != ""
    !insertmacro RefreshShortcutIfPresent "$SMPROGRAMS\$AppStartMenuFolder\${PRODUCTNAME}.lnk"
  !else
    !insertmacro RefreshShortcutIfPresent "$SMPROGRAMS\${PRODUCTNAME}.lnk"
  !endif

  !insertmacro RefreshShortcutIfPresent "$DESKTOP\${PRODUCTNAME}.lnk"
!macroend
