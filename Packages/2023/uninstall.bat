@echo off
chcp 65001 >nul
echo Revit 2023版 28 Tools をアンインストール中...

set ADDON_DIR=C:\ProgramData\Autodesk\Revit\Addins\2023

if exist "%ADDON_DIR%\28Tools" (
    rmdir /S /Q "%ADDON_DIR%\28Tools"
    echo 削除しました: 28Tools フォルダ
)

if exist "%ADDON_DIR%\Tools28.addin" (
    del "%ADDON_DIR%\Tools28.addin"
    echo 削除しました: Tools28.addin
)

REM 旧バージョンのクリーンアップ（ルートに直置きされていた場合）
if exist "%ADDON_DIR%\Tools28.dll" (
    del "%ADDON_DIR%\Tools28.dll"
    echo 削除しました: Tools28.dll（旧バージョン）
)

echo アンインストール完了！
echo Revit 2023を再起動してください。

pause
