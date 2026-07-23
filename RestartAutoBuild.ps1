# Tools28 - Restart AutoBuild
# Stops the running AutoBuild instance and starts it again in the background.
#
# Usage (from the repo root):
#   powershell -ExecutionPolicy Bypass -File .\RestartAutoBuild.ps1
#   (or right-click the file > "Run with PowerShell")
#
# AutoBuild runs elevated (it writes to C:\ProgramData\...), so this script
# self-elevates via UAC. Click "Yes" when prompted.

Set-Location $PSScriptRoot

# ----------------------------------------
# Self-elevate to administrator if needed
# ----------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Elevating to administrator (choose 'Yes' on the UAC prompt)..." -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -ArgumentList `
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`""
    return
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Tools28 - Restart AutoBuild" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------
# Stop running AutoBuild process(es)
# ----------------------------------------
Write-Host "Looking for running AutoBuild..." -ForegroundColor Gray
$procs = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like '*AutoBuild.ps1*' }

if ($procs) {
    foreach ($p in $procs) {
        Write-Host ("  Stopping PID {0}" -f $p.ProcessId) -ForegroundColor Yellow
        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
    }
    # Wait for the process to exit and release the Global\Tools28_AutoBuild mutex
    Start-Sleep -Seconds 2
    Write-Host "  Stopped." -ForegroundColor Green
} else {
    Write-Host "  No running AutoBuild found (will just start a new one)." -ForegroundColor Gray
}

Write-Host ""

# ----------------------------------------
# Start AutoBuild again (hidden, background)
# ----------------------------------------
$autoBuild = Join-Path $PSScriptRoot "AutoBuild.ps1"
if (-not (Test-Path $autoBuild)) {
    Write-Host "Error: AutoBuild.ps1 not found: $autoBuild" -ForegroundColor Red
    Start-Sleep -Seconds 4
    exit 1
}

Write-Host "Starting AutoBuild..." -ForegroundColor Gray
Start-Process powershell -WindowStyle Hidden -WorkingDirectory $PSScriptRoot -ArgumentList `
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', "`"$autoBuild`""

Start-Sleep -Seconds 1

# Confirm it came up
$running = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like '*AutoBuild.ps1*' }

Write-Host ""
if ($running) {
    Write-Host "AutoBuild restarted successfully." -ForegroundColor Green
} else {
    Write-Host "AutoBuild did not appear to start. Check the log." -ForegroundColor Yellow
}
Write-Host ("  Log: {0}" -f (Join-Path $PSScriptRoot 'AutoBuild.log')) -ForegroundColor Gray
Write-Host ""
Write-Host "This window closes in 4 seconds..." -ForegroundColor DarkGray
Start-Sleep -Seconds 4
