@echo off
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\Common7\Tools\VsDevCmd.bat" -arch=x64 -host_arch=x64 >nul
set PATH=%USERPROFILE%\.cargo\bin;%PATH%
cd /d C:\Users\PCGamer\Documents\fun
npm run dev > tauri-dev.log 2>&1
