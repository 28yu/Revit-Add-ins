@echo off
chcp 65001 >/dev/null
setlocal enabledelayedexpansion

echo Revit 2021版 28 Tools をインストール中...
echo.

REM パス設定
set ADDON_DIR=C:\ProgramData\Autodesk\Revit\Addins\2021
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

REM 28Tools フォルダごとコピー（DLL + 依存ライブラリ）
echo DLLファイルをコピー中...
if not exist "%ADDON_DIR%\28Tools" mkdir "%ADDON_DIR%\28Tools"
xcopy /Y /Q "%SCRIPT_DIR%28Tools\*.dll" "%ADDON_DIR%\28Tools\" >/dev/null
if errorlevel 1 (
    echo ✗ DLLファイルのコピーに失敗しました
    pause
    exit /b 1
)
echo ✓ DLLファイルをコピーしました

REM アドインマニフェストのコピー（Addins ルートに配置）
if exist "%SCRIPT_DIR%28Tools\Tools28.addin" (
    copy /Y "%SCRIPT_DIR%28Tools\Tools28.addin" "%ADDON_DIR%\Tools28.addin" >/dev/null
    if errorlevel 1 (
        echo ✗ Tools28.addin のコピーに失敗しました
        pause
        exit /b 1
    )
    echo ✓ Tools28.addin をコピーしました
) else (
    echo ✗ エラー: Tools28.addin が見つかりません
    pause
    exit /b 1
)

REM インストール完了確認
echo.
echo ===== インストール完了 =====
echo ✓ 以下のファイルがインストールされました:
echo   %ADDON_DIR%\Tools28.addin
dir /B "%ADDON_DIR%\28Tools\*.dll"
echo.
echo Revit 2021を完全に再起動してください。
echo リボンに「28 Tools」タブが表示されます。
echo.

pause
