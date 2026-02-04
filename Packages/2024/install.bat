@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo Revit 2024版 28 Tools をインストール中...
echo.

REM パス設定
set ADDON_DIR=C:\ProgramData\Autodesk\Revit\Addins\2024
set SCRIPT_DIR=%~dp0

REM デバッグ出力
echo ===== パス情報 =====
echo インストール先: %ADDON_DIR%
echo スクリプト場所: %SCRIPT_DIR%
echo.

REM アドインディレクトリが存在しなければ作成
if not exist "%ADDON_DIR%" (
    mkdir "%ADDON_DIR%"
    echo ✓ ディレクトリを作成しました
)

REM 28Tools フォルダの確認
if not exist "%SCRIPT_DIR%28Tools" (
    echo ✗ エラー: 28Tools フォルダが見つかりません
    echo 期待されるパス: %SCRIPT_DIR%28Tools
    pause
    exit /b 1
)

REM DLLファイルのコピー
if exist "%SCRIPT_DIR%28Tools\Tools28.dll" (
    copy /Y "%SCRIPT_DIR%28Tools\Tools28.dll" "%ADDON_DIR%\Tools28.dll"
    if errorlevel 1 (
        echo ✗ Tools28.dll のコピーに失敗しました
        pause
        exit /b 1
    )
    echo ✓ Tools28.dll をコピーしました
) else (
    echo ✗ エラー: Tools28.dll が見つかりません
    echo 期待されるパス: %SCRIPT_DIR%28Tools\Tools28.dll
    pause
    exit /b 1
)

REM アドインマニフェストのコピー
if exist "%SCRIPT_DIR%28Tools\Tools28.addin" (
    copy /Y "%SCRIPT_DIR%28Tools\Tools28.addin" "%ADDON_DIR%\Tools28.addin"
    if errorlevel 1 (
        echo ✗ Tools28.addin のコピーに失敗しました
        pause
        exit /b 1
    )
    echo ✓ Tools28.addin をコピーしました
) else (
    echo ✗ エラー: Tools28.addin が見つかりません
    echo 期待されるパス: %SCRIPT_DIR%28Tools\Tools28.addin
    pause
    exit /b 1
)

REM インストール完了確認
echo.
echo ===== インストール完了 =====
echo ✓ 以下のファイルが正常にコピーされました:
dir "%ADDON_DIR%\Tools28.*"
echo.
echo Revit 2024を完全に再起動してください。
echo リボンに「28 Tools」タブが表示されます。
echo.

pause
