@echo off
echo Revit 2023版 28 Tools をアンインストール中...

set ADDON_DIR=C:\ProgramData\Autodesk\Revit\Addins\2023

if exist "%ADDON_DIR%\28Tools" rmdir /S /Q "%ADDON_DIR%\28Tools"
if exist "%ADDON_DIR%\Tools28.addin" del "%ADDON_DIR%\Tools28.addin"
if exist "%ADDON_DIR%\Tools28.dll" del "%ADDON_DIR%\Tools28.dll"

echo アンインストール完了！
echo Revit 2023を再起動してください。

pause
