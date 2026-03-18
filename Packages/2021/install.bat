@echo off
chcp 65001 >/dev/null
setlocal enabledelayedexpansion

echo ========================================
echo   28 Tools v2.0 for Revit 2021
echo ========================================
echo.
echo インストール中...
echo.

REM パス設定
set ADDON_DIR=C:\ProgramData\Autodesk\Revit\Addins\2021
set SCRIPT_DIR=%~dp0

REM アドインディレクトリが存在しなければ作成
if not exist "%ADDON_DIR%" (
    mkdir "%ADDON_DIR%"
    echo [OK] ディレクトリを作成しました
)

REM 28Tools フォルダの確認
if not exist "%SCRIPT_DIR%28Tools" (
    echo [NG] エラー: 28Tools フォルダが見つかりません
    echo 期待されるパス: %SCRIPT_DIR%28Tools
    pause
    exit /b 1
)

REM 旧バージョンのクリーンアップ（ルートに直置きされていた場合）
if exist "%ADDON_DIR%\Tools28.dll" (
    del "%ADDON_DIR%\Tools28.dll"
    echo [OK] 旧バージョンの Tools28.dll を削除しました
)

REM 28Tools フォルダごとコピー（DLL + 依存ライブラリ）
echo ファイルをコピー中...
if not exist "%ADDON_DIR%\28Tools" mkdir "%ADDON_DIR%\28Tools"
xcopy /Y /Q "%SCRIPT_DIR%28Tools\*.dll" "%ADDON_DIR%\28Tools\" >/dev/null
if errorlevel 1 (
    echo [NG] DLLファイルのコピーに失敗しました
    pause
    exit /b 1
)
echo [OK] DLLファイルをコピーしました

REM アドインマニフェストのコピー（Addins ルートに配置）
if exist "%SCRIPT_DIR%28Tools\Tools28.addin" (
    copy /Y "%SCRIPT_DIR%28Tools\Tools28.addin" "%ADDON_DIR%\Tools28.addin" >/dev/null
    if errorlevel 1 (
        echo [NG] Tools28.addin のコピーに失敗しました
        pause
        exit /b 1
    )
    echo [OK] Tools28.addin をコピーしました
) else (
    echo [NG] エラー: Tools28.addin が見つかりません
    pause
    exit /b 1
)

echo.
echo ========================================
echo   インストール完了
echo ========================================
echo.
echo Revit 2021を完全に再起動してください。
echo リボンに「28 Tools」タブが表示されます。
echo.
echo マニュアル: https://28yu.github.io/28tools-manual/
echo.

pause
