@echo off
setlocal
cd /d "%~dp0"
set "PS1=%~dp0tools\DExcelAssist.ps1"
if not exist "%PS1%" (
  echo [ERROR] tools\DExcelAssist.ps1 not found.
  pause
  exit /b 1
)

if /i "%~1"=="/install" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Action Install
  exit /b %ERRORLEVEL%
)

if /i "%~1"=="/release" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Action Release
  exit /b %ERRORLEVEL%
)

if /i "%~1"=="/mainbranch" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Action Release
  exit /b %ERRORLEVEL%
)

echo Starting DExcelAssist v119 Installer...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
set "EC=%ERRORLEVEL%"
if not "%EC%"=="0" echo [ERROR] ExitCode=%EC%
echo.
pause
exit /b %EC%
