@echo off
rem Quick low-quality preview (about 1/8 of the normal render time).
if "%~1"=="" (
  echo Drag and drop an FBX file onto this file.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0MakePanorama.ps1" -Model "%~1" -Quality draft
pause
