@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 932 >nul

cd /d "%~dp0"
set "SCRIPT=%~dp0tools\DExcelAssist.ps1"

if not exist "%SCRIPT%" goto SCRIPT_NOT_FOUND

if /i "%~1"=="/install" goto ARG_INSTALL
if /i "%~1"=="/release" goto ARG_RELEASE
if /i "%~1"=="/mainbranch" goto ARG_RELEASE

:MENU
cls
echo ========================================
echo DExcelAssist v158 インストーラー
echo ========================================
echo.
echo 1: インストール / 修復
echo 2: 診断
echo 3: アンインストール
echo 4: 配布ファイル作成
echo 0: 終了
echo.
set "SEL="
set /p "SEL=番号を入力してください: "

if "%SEL%"=="1" goto DO_INSTALL
if "%SEL%"=="2" goto DO_DIAGNOSE
if "%SEL%"=="3" goto DO_UNINSTALL
if "%SEL%"=="4" goto DO_RELEASE
if "%SEL%"=="0" exit /b 0

echo.
echo 無効な番号です。
pause
goto MENU

:DO_INSTALL
call :RUN Install インストール / 修復
goto AFTER_RUN

:DO_DIAGNOSE
call :RUN Diagnose 診断
goto AFTER_RUN

:DO_UNINSTALL
call :RUN Uninstall アンインストール
goto AFTER_RUN

:DO_RELEASE
call :RUN Release 配布ファイル作成
goto AFTER_RUN

:AFTER_RUN
echo.
pause
goto MENU

:ARG_INSTALL
call :RUN Install インストール / 修復
exit /b %RC%

:ARG_RELEASE
call :RUN Release 配布ファイル作成
exit /b %RC%

:RUN
set "ACTION=%~1"
set "ACTION_NAME=%~2"
echo.
echo [情報] %ACTION_NAME% を実行しています...
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Action "%ACTION%"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo [エラー] 失敗しました。終了コード=%RC%
) else (
  echo [OK] 完了しました。
)
exit /b %RC%

:SCRIPT_NOT_FOUND
echo [エラー] tools\DExcelAssist.ps1 が見つかりません。
echo Path: %SCRIPT%
pause
exit /b 1
