@echo off
rem Drag and drop an FBX file exported from Revit onto this file.
rem Renders a 360-degree panorama with Blender and creates an HTML viewer.
if "%~1"=="" (
  echo Drag and drop an FBX file onto this file.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0MakePanorama.ps1" -Model "%~1" -Quality normal
pause
