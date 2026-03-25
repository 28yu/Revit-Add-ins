# Tools28 - AutoBuild Auto-Start Setup
# Usage: Run once. No admin rights needed.
# Uninstall: .\SetupAutoStart.ps1 -Remove

param(
    [switch]$Remove
)

$startupFolder = [Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startupFolder "Tools28 AutoBuild.lnk"

if ($Remove) {
    if (Test-Path $shortcutPath) {
        Remove-Item $shortcutPath -Force
        Write-Host "Auto-start removed." -ForegroundColor Green
        Write-Host "  Deleted: $shortcutPath" -ForegroundColor Gray
    } else {
        Write-Host "Auto-start is not configured." -ForegroundColor Yellow
    }
    return
}

$repoRoot = $PSScriptRoot
$autoBuildScript = Join-Path $repoRoot "AutoBuild.ps1"

if (-not (Test-Path $autoBuildScript)) {
    Write-Host "Error: AutoBuild.ps1 not found: $autoBuildScript" -ForegroundColor Red
    exit 1
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$autoBuildScript`""
$shortcut.WorkingDirectory = $repoRoot
$shortcut.Description = "Tools28 AutoBuild"
$shortcut.WindowStyle = 1
$shortcut.Save()

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  AutoBuild auto-start configured!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Shortcut: $shortcutPath" -ForegroundColor Gray
Write-Host "  Script:   $autoBuildScript" -ForegroundColor Gray
Write-Host ""
Write-Host "  AutoBuild will start automatically on next Windows logon." -ForegroundColor Green
Write-Host ""
Write-Host "  To remove: .\SetupAutoStart.ps1 -Remove" -ForegroundColor Yellow
Write-Host ""
