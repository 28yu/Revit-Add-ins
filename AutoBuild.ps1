# Tools28 - Auto Build & Deploy Monitor (v17)
# Usage:
#   .\AutoBuild.ps1                     # Default (30s interval)
#   .\AutoBuild.ps1 -Interval 60        # 60s interval
#
# Monitors origin/main for changes, auto pulls, builds, and deploys.
# Press Ctrl+C to stop.

param(
    [Parameter(Mandatory=$false)]
    [int]$Interval = 30
)

# ========================================
# Ensure working directory is repo root
# ========================================
Set-Location $PSScriptRoot

# ========================================
# Setup
# ========================================

$LogFile = Join-Path $PSScriptRoot "AutoBuild.log"

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $Message"
    Write-Host $logLine -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $logLine -ErrorAction SilentlyContinue
}

function Show-Notification {
    param([string]$Body, [string]$Title, [string]$Icon = "Information")
    Start-Process powershell -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-Command', "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('$Body', '$Title', 'OK', '$Icon')" -WindowStyle Hidden
}

function Get-RevitVersion {
    $configPath = ".\dev-config.json"
    $ver = "2022"
    if (Test-Path $configPath) {
        try {
            $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
            $ver = $cfg.defaultRevitVersion
        } catch { }
    }
    return $ver
}

function Run-Build {
    # Run QuickBuild and check success by DLL timestamp (exit code is unreliable via & .\script.ps1)
    $revitVer = Get-RevitVersion
    $dllPath = ".\bin\Release\Revit$revitVer\Tools28.dll"

    # Record DLL timestamp before build
    $dllTimeBefore = $null
    if (Test-Path $dllPath) {
        $dllTimeBefore = (Get-Item $dllPath).LastWriteTime
    }

    # Capture build output to log file for diagnostics
    $buildLog = Join-Path $PSScriptRoot "AutoBuild_detail.log"
    $output = & .\QuickBuild.ps1 2>&1
    $exitCode = $LASTEXITCODE
    $output | Out-File -FilePath $buildLog -Encoding UTF8 -Force

    # Check success: DLL exists AND was updated (newer timestamp)
    $dllExists = Test-Path $dllPath
    $dllUpdated = $false
    if ($dllExists) {
        $dllTimeAfter = (Get-Item $dllPath).LastWriteTime
        $dllUpdated = ($dllTimeBefore -eq $null) -or ($dllTimeAfter -gt $dllTimeBefore)
    }
    $success = $dllExists -and $dllUpdated

    Write-Log "Build result: exitCode=$exitCode, dllExists=$dllExists, dllUpdated=$dllUpdated, success=$success" "Gray"

    # Log build details on failure
    if (-not $success) {
        Write-Log "--- Build output (last 20 lines) ---" "Yellow"
        if (Test-Path $buildLog) {
            Get-Content $buildLog -Tail 20 | ForEach-Object { Write-Log "  $_" "Gray" }
        } else {
            Write-Log "  (no build log created)" "Red"
        }
        Write-Log "--- End build output ---" "Yellow"
    }

    return $success
}

# ========================================
# Duplicate instance prevention
# ========================================
$mutex = New-Object System.Threading.Mutex($false, "Global\Tools28_AutoBuild")
if (-not $mutex.WaitOne(0)) {
    # Another instance is already running - exit silently
    exit 0
}

$host.UI.RawUI.WindowTitle = "Tools28 AutoBuild"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Tools28 - Auto Build & Deploy" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Interval: ${Interval}s" -ForegroundColor Gray
Write-Host "  Stop: Ctrl+C" -ForegroundColor Gray
Write-Host ""

# Verify we are in the repo root
if (-not (Test-Path ".\Tools28.csproj")) {
    Write-Host "Error: Tools28.csproj not found. Run from repo root." -ForegroundColor Red
    exit 1
}

# Switch to main branch if needed
$currentBranch = git rev-parse --abbrev-ref HEAD 2>$null
if ($currentBranch -ne "main") {
    Write-Host "Switching to main branch..." -ForegroundColor Yellow
    git checkout main 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to switch to main branch." -ForegroundColor Red
        exit 1
    }
}

# Initial fetch
git fetch origin main 2>$null
$remoteLatest = git rev-parse origin/main 2>$null
$localHead = git rev-parse HEAD 2>$null

Write-Log "Started (local: $($localHead.Substring(0, 7)), remote: $($remoteLatest.Substring(0, 7)))" "Green"

# ========================================
# Startup check: build immediately if behind
# ========================================

if ($localHead -ne $remoteLatest) {
    Write-Host ""
    Write-Log "Local is behind remote. Building now..." "Yellow"
    Write-Host ""

    git reset --hard origin/main 2>$null
    git clean -fd 2>$null

    $localHead = git rev-parse HEAD 2>$null
    $commitMsg = git log HEAD -1 --format="%s" 2>$null
    Write-Log "Pull OK (HEAD: $($localHead.Substring(0,7)))" "Green"
    Write-Log "Building..." "Yellow"
    Write-Host ""

    $buildSuccess = Run-Build
    $shortInfo = "$commitMsg ($($localHead.Substring(0,7)))"

    if ($buildSuccess) {
        Write-Log "Startup build OK! Restart Revit to test." "Green"
        Show-Notification "Deploy OK - $shortInfo" "Tools28 Build OK" "Information"
    } else {
        Write-Log "Startup build FAILED." "Red"
        Show-Notification "Build FAILED - $shortInfo" "Tools28 Build FAILED" "Error"
    }
    Write-Host ""
} else {
    Write-Log "Local is up to date. Waiting for changes..." "Green"
}

# Re-fetch before entering loop (prevent double-build if another merge happened during build)
git fetch origin main 2>$null
$lastCommit = git rev-parse origin/main 2>$null
Write-Host ""

# ========================================
# Monitor loop
# ========================================

$buildCount = 0

while ($true) {
    try {
        git fetch origin main 2>$null

        if ($LASTEXITCODE -ne 0) {
            Write-Log "Fetch failed (check network)" "Yellow"
            Start-Sleep -Seconds $Interval
            continue
        }

        $remoteCommit = git rev-parse origin/main 2>$null

        if ($remoteCommit -ne $lastCommit) {
            $buildCount++
            $commitMsg = git log origin/main -1 --format="%s" 2>$null

            Write-Host ""
            Write-Log "Change detected! (Build #$buildCount) $commitMsg" "Yellow"
            Write-Host ""

            Write-Log "Pulling..." "Yellow"
            git reset --hard origin/main 2>$null
            git clean -fd 2>$null

            $localHead = git rev-parse HEAD 2>$null
            if ($localHead -ne $remoteCommit) {
                Write-Log "Error: HEAD mismatch after pull (local=$($localHead.Substring(0,7)) remote=$($remoteCommit.Substring(0,7)))" "Red"
                Show-Notification "Git pull failed" "Tools28 Build FAILED" "Error"
                $lastCommit = $remoteCommit
                Start-Sleep -Seconds $Interval
                continue
            }

            Write-Log "Pull OK (HEAD: $($localHead.Substring(0,7)))" "Green"
            Write-Log "Building..." "Yellow"
            Write-Host ""

            $buildSuccess = Run-Build
            $shortHash = $remoteCommit.Substring(0, 7)
            $shortInfo = "$commitMsg ($shortHash)"

            if ($buildSuccess) {
                Write-Log "Build & Deploy OK! Restart Revit to test." "Green"
                Show-Notification "Deploy OK - $shortInfo" "Tools28 Build OK" "Information"
            } else {
                Write-Log "Build FAILED." "Red"
                Show-Notification "Build FAILED - $shortInfo" "Tools28 Build FAILED" "Error"
            }

            $lastCommit = $remoteCommit
            Write-Host ""
            Write-Host "Monitoring..." -ForegroundColor Gray
        } else {
            Write-Host "." -NoNewline -ForegroundColor DarkGray
        }

        Start-Sleep -Seconds $Interval
    }
    catch {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host ""
        Write-Host "[$timestamp] Error: $_" -ForegroundColor Red
        Start-Sleep -Seconds $Interval
    }
}
