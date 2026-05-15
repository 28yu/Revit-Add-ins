# Formwork diagnostic log uploader for Claude Code
#
# Usage:
#   cd C:\actions-runner\_work\Revit-Add-ins\Revit-Add-ins
#   git pull origin main
#   .\Send-FormworkLog.ps1
#
# Copies Formwork_debug.txt AND any timestamped backup logs (Formwork_debug_*.txt)
# from C:\temp\ to .diag\ and pushes to git so Claude can read them.

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$srcDir   = "C:\temp"
$dstDir   = Join-Path $repoRoot ".diag"

# --- collect all Formwork log files from C:\temp ---
$logFiles = @()

$main = Join-Path $srcDir "Formwork_debug.txt"
if (Test-Path $main) { $logFiles += $main }

$backups = Get-ChildItem -Path $srcDir -Filter "Formwork_debug_*.txt" -ErrorAction SilentlyContinue
foreach ($f in $backups) { $logFiles += $f.FullName }

if ($logFiles.Count -eq 0) {
    Write-Host "ERROR: No Formwork_debug*.txt found in $srcDir" -ForegroundColor Red
    Write-Host "  Run formwork takeoff in Revit first." -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $dstDir)) {
    New-Item -ItemType Directory -Path $dstDir | Out-Null
}

# --- copy each file ---
$copied = @()
foreach ($src in $logFiles) {
    $name = Split-Path $src -Leaf
    $dst  = Join-Path $dstDir $name
    Copy-Item $src $dst -Force
    $sizeKB = [math]::Round((Get-Item $dst).Length / 1024, 1)
    Write-Host ("  Copied: {0} ({1} KB)" -f $name, $sizeKB) -ForegroundColor Cyan
    $copied += ".diag\$name"
}

Write-Host ("{0} file(s) copied." -f $copied.Count) -ForegroundColor Green

# --- git add / commit / push ---
Push-Location $repoRoot
try {
    foreach ($rel in $copied) {
        git add -f $rel
    }
    git commit -m "diag: Formwork_debug logs ($($copied.Count) files)"
    git push
    Write-Host "Push done. Tell Claude that the log has been updated." -ForegroundColor Green
}
catch {
    Write-Host ("Git error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
