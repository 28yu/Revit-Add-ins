# Formwork diagnostic log uploader for Claude Code
#
# Usage:
#   1. Run formwork takeoff in Revit -> C:\temp\Formwork_debug.txt is updated
#   2. Run this script: .\Send-FormworkLog.ps1
#   3. Tell Claude that the log has been updated
#
# Note: ASCII only to avoid PowerShell 5.1 UTF-8/Shift-JIS encoding issues.

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$srcLog   = "C:\temp\Formwork_debug.txt"
$dstDir   = Join-Path $repoRoot ".diag"
$dstLog   = Join-Path $dstDir   "Formwork_debug.txt"

if (-not (Test-Path $srcLog)) {
    Write-Host "ERROR: $srcLog not found. Run formwork takeoff in Revit first." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $dstDir)) {
    New-Item -ItemType Directory -Path $dstDir | Out-Null
}

Copy-Item $srcLog $dstLog -Force

$size = (Get-Item $dstLog).Length
$sizeKB = [math]::Round($size / 1024, 1)
Write-Host ("Copied: {0} ({1} KB)" -f $dstLog, $sizeKB) -ForegroundColor Green

Push-Location $repoRoot
try {
    git add -f .\.diag\Formwork_debug.txt
    git commit -m "diag: update Formwork_debug.txt"
    git push
    Write-Host "Push done. Tell Claude that the log has been updated." -ForegroundColor Green
}
finally {
    Pop-Location
}
