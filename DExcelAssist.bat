@echo off


setlocal


cd /d "%~dp0"


echo Starting DExcelAssist v112 Unified Ribbon...


echo.


set "PS1=%~dp0tools\DExcelAssist.ps1"


if not exist "%PS1%" (


  echo [ERROR] tools\DExcelAssist.ps1 not found.


  pause


  exit /b 1


)


powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%"


set "EC=%ERRORLEVEL%"


if not "%EC%"=="0" echo [ERROR] ExitCode=%EC%


echo.


pause


exit /b %EC%


