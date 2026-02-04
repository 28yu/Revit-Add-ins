@echo off
chcp 65001 >nul
echo Revit 2026版 28 Tools をアンインストール中...

set ADDON_DIR=C:\ProgramData\Autodesk\Revit\Addins\2026

if exist "%ADDON_DIR%\Tools28.dll" (
    del "%ADDON_DIR%\Tools28.dll"
    echo 削除しました: Tools28.dll
)

if exist "%ADDON_DIR%\Tools28.addin" (
    del "%ADDON_DIR%\Tools28.addin"
    echo 削除しました: Tools28.addin
)

echo アンインストール完了！
echo Revit 2026を再起動してください。

pause
