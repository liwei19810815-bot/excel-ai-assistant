@echo off
setlocal
cd /d "%~dp0.."

rem Dev sideload only. Neither a plain "register" nor a normal
rem double-click Excel launch will make the ribbon button appear
rem (verified on real hardware) -- Office only injects the ribbon
rem customization through this explicit sideload launch. --no-debug
rem skips the "Web view stopped while loading" debugger-attach dialog.
rem Comments kept ASCII-only: Chinese text in a BOM-less .bat gets
rem misdecoded by cmd's default codepage and corrupts later lines.

echo Checking dev server (localhost:3000)...
powershell -NoProfile -Command "try { (Invoke-WebRequest -Uri https://localhost:3000/manifest.xml -UseBasicParsing -TimeoutSec 2).StatusCode } catch { 0 }" > "%TEMP%\_excelai_devcheck.txt"
set /p DEVSTATUS=<"%TEMP%\_excelai_devcheck.txt"
del "%TEMP%\_excelai_devcheck.txt" >nul 2>&1

if not "%DEVSTATUS%"=="200" (
  echo Dev server not running, starting it...
  start "Excel AI dev server" cmd /k "cd /d %~dp0.. && npm run dev"
  echo Waiting for server to come up...
  timeout /t 6 /nobreak >nul
)

echo Launching Excel with the AI add-in...
call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug

endlocal
