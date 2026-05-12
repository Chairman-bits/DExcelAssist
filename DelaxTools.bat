@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem ==================================================
rem DelaxTools Installer v166
rem - 文字化け・改行形式による即終了を避けるため、内部処理を単純化
rem - エラー時も画面を閉じず、必ず停止して確認できるようにする
rem ==================================================

chcp 932 >nul
title DelaxTools v166 インストーラー

pushd "%~dp0" >nul 2>&1
if errorlevel 1 goto ROOT_ERROR

set "SCRIPT=%~dp0tools\DelaxTools.ps1"
if not exist "%SCRIPT%" goto SCRIPT_NOT_FOUND

if /i "%~1"=="/install" goto ARG_INSTALL
if /i "%~1"=="/release" goto ARG_RELEASE
if /i "%~1"=="/mainbranch" goto ARG_RELEASE

:MENU
cls
echo ========================================
echo DelaxTools v166 インストーラー
echo ========================================
echo.
echo 1: インストール / 修復
echo 2: 診断
echo 3: アンインストール
echo 4: 配布ファイル作成
echo 0: 終了
echo.
set "SEL="
set /p "SEL=番号を選択してください: "

if "%SEL%"=="1" goto DO_INSTALL
if "%SEL%"=="2" goto DO_DIAGNOSE
if "%SEL%"=="3" goto DO_UNINSTALL
if "%SEL%"=="4" goto DO_RELEASE
if "%SEL%"=="0" goto END_OK

echo.
echo 正しい番号を入力してください。
pause
goto MENU

:DO_INSTALL
set "ACTION=Install"
set "ACTION_NAME=インストール / 修復"
goto RUN

:DO_DIAGNOSE
set "ACTION=Diagnose"
set "ACTION_NAME=診断"
goto RUN

:DO_UNINSTALL
set "ACTION=Uninstall"
set "ACTION_NAME=アンインストール"
goto RUN

:DO_RELEASE
set "ACTION=Release"
set "ACTION_NAME=配布ファイル作成"
goto RUN

:ARG_INSTALL
set "ACTION=Install"
set "ACTION_NAME=インストール / 修復"
goto RUN_ARG

:ARG_RELEASE
set "ACTION=Release"
set "ACTION_NAME=配布ファイル作成"
goto RUN_ARG

:RUN
echo.
echo [実行] %ACTION_NAME% を開始します...
call :CALL_POWERSHELL
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
  echo [エラー] 処理が失敗しました。終了コード=%RC%
) else (
  echo [OK] 処理が完了しました。
)
echo.
pause
goto MENU

:RUN_ARG
echo.
echo [実行] %ACTION_NAME% を開始します...
call :CALL_POWERSHELL
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
  echo [エラー] 処理が失敗しました。終了コード=%RC%
) else (
  echo [OK] 処理が完了しました。
)
echo.
pause
exit /b %RC%

:CALL_POWERSHELL
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Action "%ACTION%"
exit /b %ERRORLEVEL%

:SCRIPT_NOT_FOUND
echo.
echo [エラー] tools\DelaxTools.ps1 が見つかりません。
echo Path: %SCRIPT%
echo.
pause
exit /b 1

:ROOT_ERROR
echo.
echo [エラー] インストーラーのフォルダへ移動できませんでした。
echo Path: %~dp0
echo.
pause
exit /b 1

:END_OK
popd >nul 2>&1
exit /b 0
