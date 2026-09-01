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

# Shared deploy helpers (handles DLLs locked by a running Revit)
. (Join-Path $PSScriptRoot "DeployHelpers.ps1")

# ========================================
# Setup
# ========================================

# Git output is UTF-8, but PowerShell 5.1 defaults to system locale (Shift-JIS)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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
    # Write JSON with .NET to ensure proper UTF-8 encoding
    $dataFile = Join-Path $env:TEMP "Tools28_notify.json"
    $json = @{ Body = $Body; Title = $Title; Icon = $Icon } | ConvertTo-Json
    [System.IO.File]::WriteAllText($dataFile, $json, [System.Text.Encoding]::UTF8)
    # Read script uses .NET for reliable UTF-8 reading
    $readScript = "`$json = [System.IO.File]::ReadAllText('$dataFile', [System.Text.Encoding]::UTF8); `$d = `$json | ConvertFrom-Json; Add-Type -AssemblyName System.Windows.Forms; `$i = if (`$d.Icon -eq 'Error') {[System.Windows.Forms.MessageBoxIcon]::Error} else {[System.Windows.Forms.MessageBoxIcon]::Information}; [System.Windows.Forms.MessageBox]::Show(`$d.Body, `$d.Title, [System.Windows.Forms.MessageBoxButtons]::OK, `$i); Remove-Item '$dataFile' -Force -ErrorAction SilentlyContinue"
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($readScript)
    $encoded = [Convert]::ToBase64String($bytes)
    Start-Process powershell -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-EncodedCommand', $encoded -WindowStyle Hidden
}

# ----------------------------------------
# Japanese notification texts
# ----------------------------------------
# PowerShell 5.1 reads .ps1 files using the system locale (Shift-JIS), so
# Japanese literals inside this script can turn into mojibake.
# The texts therefore live in AutoBuild.messages.json (UTF-8) and are read
# with an explicit UTF-8 decoder.
function Get-Msg {
    param([string]$Key, [string[]]$FormatArgs = @())

    # $PSScriptRoot is empty when the function is dot-sourced from a console,
    # so fall back to the current directory (AutoBuild sets it to the repo root)
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }

    $messages = $null
    try {
        $msgPath = Join-Path $scriptDir "AutoBuild.messages.json"
        $raw = [System.IO.File]::ReadAllText($msgPath, [System.Text.Encoding]::UTF8)
        $messages = $raw | ConvertFrom-Json
    } catch { }

    $text = $null
    if ($messages) { $text = $messages.$Key }
    # Fall back to the key name so a missing entry is visible instead of blank
    if (-not $text) { return $Key }

    # Cast to object[] so PowerShell picks Format(string, params object[])
    # instead of Format(string, object) (which would print "System.String[]")
    if ($FormatArgs.Count -gt 0) { return [string]::Format($text, [object[]]$FormatArgs) }
    return $text
}

$AllRevitVersions = @("2021", "2022", "2023", "2024", "2025", "2026")

# Versions built by the most recent Run-Build call (for notifications)
$script:LastBuildVersions = @()

# How the most recent deploy landed (for notifications):
#   APPLIED          - copied straight in (Revit was not running)
#   RESTART_REQUIRED - Revit was running; new DLLs are in place, restart to pick them up
#   PENDING          - could not swap the files; they wait in the pending folder
#   FAILED           - deploy failed
$script:LastDeployStatus = "UNKNOWN"

function Get-DefaultRevitVersion {
    $configPath = ".\dev-config.json"
    $ver = "2022"
    if (Test-Path $configPath) {
        try {
            $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
            if ($cfg.defaultRevitVersion) { $ver = "$($cfg.defaultRevitVersion)" }
        } catch { }
    }
    return $ver
}

function Get-BuildVersions {
    # Determine which Revit version(s) to build for the current HEAD commit.
    #
    # If the commit message contains a marker, use those versions for THIS build only:
    #   [build:2024]          -> Revit 2024
    #   [build:2024,2025]     -> Revit 2024 and 2025
    #   [build:all]           -> all supported versions (2021-2026)
    # Otherwise, fall back to the default from dev-config.json (normally 2022).
    #
    # NOTE: The claude/** auto-merge workflow squashes to main keeping only the
    # FIRST LINE of the commit, so the marker must be placed in the commit SUBJECT.
    $msg = (git log HEAD -1 --format="%B" 2>$null) -join "`n"
    if ($msg) {
        $m = [regex]::Match($msg, '\[build:\s*([0-9,\s]+|all)\s*\]',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) {
            $raw = $m.Groups[1].Value.Trim()
            if ($raw -match '(?i)all') {
                return ,@($AllRevitVersions)
            }
            $vers = $raw -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object { $AllRevitVersions -contains $_ } |
                Select-Object -Unique
            if ($vers.Count -gt 0) {
                return ,@($vers)
            }
        }
    }
    return ,@((Get-DefaultRevitVersion))
}

function Run-Build {
    # Run QuickBuild for each target version and check success by DLL timestamp
    # (exit code is unreliable via & .\script.ps1).
    $versions = Get-BuildVersions
    $script:LastBuildVersions = $versions
    Write-Log "Build target version(s): $($versions -join ', ')" "Cyan"

    # Fresh build log for this run (per-version output is appended below)
    $buildLog = Join-Path $PSScriptRoot "AutoBuild_detail.log"
    Set-Content -Path $buildLog -Value "" -Encoding UTF8 -ErrorAction SilentlyContinue

    $allSuccess = $true

    # Worst status across all built versions (FAILED > PENDING > RESTART_REQUIRED > APPLIED)
    $statusRank = @{ "APPLIED" = 0; "RESTART_REQUIRED" = 1; "PENDING" = 2; "UNKNOWN" = 3; "FAILED" = 4 }
    $worstStatus = "APPLIED"

    foreach ($revitVer in $versions) {
        $dllPath = ".\bin\Release\Revit$revitVer\Tools28.dll"

        # Record DLL timestamp before build
        $dllTimeBefore = $null
        if (Test-Path $dllPath) {
            $dllTimeBefore = (Get-Item $dllPath).LastWriteTime
        }

        # Build the explicitly-requested version (do NOT rely on dev-config default)
        Add-Content -Path $buildLog -Value "===== Revit $revitVer =====" -Encoding UTF8 -ErrorAction SilentlyContinue
        # *>&1 captures every stream (including Write-Host / information),
        # so AutoBuild_detail.log holds the full deploy output as well.
        $output = & .\QuickBuild.ps1 -RevitVersion $revitVer *>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -FilePath $buildLog -Encoding UTF8 -Append

        # QuickBuild.ps1 prints "DEPLOY_STATUS=<state>" as its last word on the deploy
        $statusLine = $output |
            ForEach-Object { "$_" } |
            Select-String -Pattern '^DEPLOY_STATUS=(\w+)' |
            Select-Object -Last 1
        $deployStatus = if ($statusLine) { $statusLine.Matches[0].Groups[1].Value } else { "UNKNOWN" }

        if ($statusRank[$deployStatus] -gt $statusRank[$worstStatus]) { $worstStatus = $deployStatus }

        $dllExists = Test-Path $dllPath
        $dllUpdated = $false
        if ($dllExists) {
            $dllTimeAfter = (Get-Item $dllPath).LastWriteTime
            $dllUpdated = ($dllTimeBefore -eq $null) -or ($dllTimeAfter -gt $dllTimeBefore)
        }

        # Success: the deploy reported a state where the new DLL will reach Revit.
        # PENDING counts as success - the build itself is fine and the files are
        # applied automatically once Revit is closed.
        if ($deployStatus -eq "APPLIED" -or $deployStatus -eq "RESTART_REQUIRED" -or $deployStatus -eq "PENDING") {
            $success = $dllExists
        } elseif ($deployStatus -eq "UNKNOWN") {
            # Older fallback: compare the DLL timestamp before/after the build
            $success = $dllExists -and $dllUpdated
        } else {
            $success = $false
        }

        Write-Log "Build[$revitVer] result: exitCode=$exitCode, dllExists=$dllExists, dllUpdated=$dllUpdated, deploy=$deployStatus, success=$success" "Gray"

        # Log build details on failure
        if (-not $success) {
            $allSuccess = $false
            Write-Log "--- Build[$revitVer] output (last 20 lines) ---" "Yellow"
            if (Test-Path $buildLog) {
                Get-Content $buildLog -Tail 20 | ForEach-Object { Write-Log "  $_" "Gray" }
            } else {
                Write-Log "  (no build log created)" "Red"
            }
            Write-Log "--- End build output ---" "Yellow"
        }
    }

    $script:LastDeployStatus = $worstStatus
    Write-Log "Deploy status: $worstStatus" "Gray"

    return $allSuccess
}

# ----------------------------------------
# Build result notification
# ----------------------------------------
function Show-BuildResultNotification {
    param([bool]$Success, [string]$CommitMsg)

    $versionsText = ($script:LastBuildVersions -join ', ')

    if ($Success) {
        switch ($script:LastDeployStatus) {
            "APPLIED"          { $detail = Get-Msg "DeployApplied"         @($versionsText) }
            "RESTART_REQUIRED" { $detail = Get-Msg "DeployRestartRequired" @($versionsText) }
            "PENDING"          { $detail = Get-Msg "DeployPending"         @($versionsText) }
            default            { $detail = Get-Msg "DeployApplied"         @($versionsText) }
        }
        Show-Notification "$CommitMsg`n`n$detail" (Get-Msg "SuccessTitle") "Information"
    }
    elseif ($script:LastDeployStatus -eq "FAILED") {
        # Compiled fine, but the files could not be placed
        $detail = Get-Msg "DeployFailed" @($versionsText)
        Show-Notification "$CommitMsg`n`n$detail" (Get-Msg "FailedTitle") "Error"
    }
    else {
        Show-Notification ((Get-Msg "BuildFailed") + "`n`n$CommitMsg") (Get-Msg "FailedTitle") "Error"
    }
}

# ----------------------------------------
# Apply files that were left pending while Revit was running
# ----------------------------------------
# Called on every idle tick of the monitor loop. Once Revit is closed the
# staged files are copied in and the *.old backups are cleaned up, so the
# next Revit start picks up the latest build.
function Invoke-PendingDeployIfPossible {
    if (Test-RevitRunning) { return }

    $appliedVersions = @()
    foreach ($ver in $AllRevitVersions) {
        $targetDir = Get-Tools28TargetDir $ver
        if (-not (Test-Path $targetDir)) { continue }

        $result = Invoke-Tools28PendingDeploy -RevitVersion $ver
        if ($result.Applied -gt 0) {
            Write-Host ""
            Write-Log "Applied $($result.Applied) pending file(s) for Revit $ver" "Green"
            $appliedVersions += $ver
        }
    }

    if ($appliedVersions.Count -gt 0) {
        Show-Notification (Get-Msg "PendingAppliedBody" @(($appliedVersions -join ', '))) `
            (Get-Msg "PendingAppliedTitle") "Information"
    }
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
        Write-Log "Startup build OK (deploy: $($script:LastDeployStatus))" "Green"
    } else {
        Write-Log "Startup build FAILED (deploy: $($script:LastDeployStatus))" "Red"
    }
    Show-BuildResultNotification $buildSuccess $commitMsg
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
                Write-Log "Build & Deploy OK (deploy: $($script:LastDeployStatus))" "Green"
            } else {
                Write-Log "Build FAILED (deploy: $($script:LastDeployStatus))" "Red"
            }
            Show-BuildResultNotification $buildSuccess $commitMsg

            $lastCommit = $remoteCommit
            Write-Host ""
            Write-Host "Monitoring..." -ForegroundColor Gray
        } else {
            # No new commit: if Revit has been closed, flush anything that was
            # left pending while it was running.
            Invoke-PendingDeployIfPossible
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
