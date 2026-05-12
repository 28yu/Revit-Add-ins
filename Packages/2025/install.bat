@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   28 Tools v2.1 for Revit 2025
echo ========================================
echo.
echo インストール中...
echo.

REM Revit 起動中チェック (DLL ファイルロック回避)
tasklist /FI "IMAGENAME eq Revit.exe" 2>/dev/null | find /I "Revit.exe" >/dev/null
if errorlevel 1 goto revit_ok
echo [NG] Revit が起動しています。
echo Revit を終了してから再度このスクリプトを実行してください。
echo.
pause
exit /b 1
:revit_ok

REM パス設定
set ADDON_DIR=C:\ProgramData\Autodesk\Revit\Addins\2025
set SCRIPT_DIR=%~dp0

REM アドインディレクトリが存在しなければ作成
if not exist "%ADDON_DIR%" (
    mkdir "%ADDON_DIR%"
    echo [OK] ディレクトリを作成しました
)

REM 28Tools フォルダの確認
if not exist "%SCRIPT_DIR%28Tools" goto no_folder
goto folder_ok
:no_folder
echo [NG] エラー: 28Tools フォルダが見つかりません
echo 期待されるパス: %SCRIPT_DIR%28Tools
pause
exit /b 1
:folder_ok

REM 旧バージョンのクリーンアップ（ルートに直置きされていた場合）
if exist "%ADDON_DIR%\Tools28.dll" (
    del "%ADDON_DIR%\Tools28.dll"
    echo [OK] 旧バージョンの Tools28.dll を削除しました
)

REM 28Tools フォルダごとコピー（DLL + 依存ライブラリ）
echo ファイルをコピー中...
if not exist "%ADDON_DIR%\28Tools" mkdir "%ADDON_DIR%\28Tools"
xcopy /Y /Q "%SCRIPT_DIR%28Tools\*.dll" "%ADDON_DIR%\28Tools\" >/dev/null
if errorlevel 1 goto copy_fail
goto copy_ok
:copy_fail
echo [NG] DLLファイルのコピーに失敗しました
echo Revit が起動中、または管理者権限が必要な可能性があります。
echo Revit を終了し、install.bat を右クリック ^> 管理者として実行してください。
pause
exit /b 1
:copy_ok
echo [OK] DLLファイルをコピーしました

REM アドインマニフェストのコピー（Addins ルートに配置）
if not exist "%SCRIPT_DIR%28Tools\Tools28.addin" goto no_addin
copy /Y "%SCRIPT_DIR%28Tools\Tools28.addin" "%ADDON_DIR%\Tools28.addin" >/dev/null
if errorlevel 1 goto addin_fail
echo [OK] Tools28.addin をコピーしました
goto addin_ok
:no_addin
echo [NG] エラー: Tools28.addin が見つかりません
pause
exit /b 1
:addin_fail
echo [NG] Tools28.addin のコピーに失敗しました
pause
exit /b 1
:addin_ok

echo.
echo ========================================
echo   インストール完了
echo ========================================
echo.
echo Revit 2025を完全に再起動してください。
echo リボンに「28 Tools」タブが表示されます。
echo.
echo マニュアル: https://28tools.com/addins.html
echo.

pause
