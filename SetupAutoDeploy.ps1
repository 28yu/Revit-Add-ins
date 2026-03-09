# Tools28 - 自動デプロイ環境セットアップ
# 使用方法:
#   .\SetupAutoDeploy.ps1            # AutoBuild監視をタスクスケジューラに登録
#   .\SetupAutoDeploy.ps1 -Remove    # タスクの登録解除
#
# 概要:
#   AutoBuild.ps1 を Windows タスクスケジューラに登録し、
#   ログオン時に自動的に main ブランチの監視を開始します。
#
# パイプライン:
#   claude/* push → auto-merge (GitHub Actions) → main 更新
#   → AutoBuild.ps1 が検知 → pull → ビルド → デプロイ → 通知

param(
    [switch]$Remove
)

$TaskName = "Tools28_AutoBuild"
$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Tools28 - 自動デプロイ セットアップ" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# 登録解除モード
# ========================================
if ($Remove) {
    Write-Host "タスク「$TaskName」を削除します..." -ForegroundColor Yellow

    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "タスクを削除しました。" -ForegroundColor Green
    } else {
        Write-Host "タスクが見つかりません（既に削除済み）。" -ForegroundColor Gray
    }

    # スタートアップショートカットも削除
    $startupLink = Join-Path ([Environment]::GetFolderPath("Startup")) "Tools28_AutoBuild.lnk"
    if (Test-Path $startupLink) {
        Remove-Item $startupLink -Force
        Write-Host "スタートアップショートカットを削除しました。" -ForegroundColor Green
    }

    Write-Host ""
    exit 0
}

# ========================================
# 前提条件チェック
# ========================================

$autoBuildScript = Join-Path $ScriptRoot "AutoBuild.ps1"
if (-not (Test-Path $autoBuildScript)) {
    Write-Host "エラー: AutoBuild.ps1 が見つかりません: $autoBuildScript" -ForegroundColor Red
    exit 1
}

$csproj = Join-Path $ScriptRoot "Tools28.csproj"
if (-not (Test-Path $csproj)) {
    Write-Host "エラー: Tools28.csproj が見つかりません。" -ForegroundColor Red
    Write-Host "リポジトリのルートディレクトリで実行してください。" -ForegroundColor Yellow
    exit 1
}

# git が使えるか確認
$gitVersion = git --version 2>$null
if (-not $gitVersion) {
    Write-Host "エラー: git が見つかりません。" -ForegroundColor Red
    exit 1
}

Write-Host "リポジトリ: $ScriptRoot" -ForegroundColor Gray
Write-Host ""

# ========================================
# 方法選択
# ========================================

Write-Host "自動デプロイの方法を選択してください:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  [1] タスクスケジューラ登録 (推奨)" -ForegroundColor White
Write-Host "      → ログオン時に自動で AutoBuild.ps1 監視が起動" -ForegroundColor Gray
Write-Host ""
Write-Host "  [2] スタートアップフォルダにショートカット作成" -ForegroundColor White
Write-Host "      → ログオン時にターミナルウィンドウが開いて監視開始" -ForegroundColor Gray
Write-Host ""

$choice = Read-Host "選択 (1 or 2)"

switch ($choice) {
    "1" {
        # ========================================
        # 方法1: タスクスケジューラ登録
        # ========================================

        Write-Host ""
        Write-Host "タスクスケジューラに登録します..." -ForegroundColor Yellow

        # 既存タスクの削除
        $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "既存のタスクを更新します..." -ForegroundColor Gray
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        }

        # PowerShell コマンド構築
        $pwshExe = "powershell.exe"
        if (Get-Command "pwsh.exe" -ErrorAction SilentlyContinue) {
            $pwshExe = "pwsh.exe"
        }

        $arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -File `"$autoBuildScript`""

        # タスク作成
        $action = New-ScheduledTaskAction `
            -Execute $pwshExe `
            -Argument $arguments `
            -WorkingDirectory $ScriptRoot

        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit ([TimeSpan]::Zero)

        Register-ScheduledTask `
            -TaskName $TaskName `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Description "Tools28 自動ビルド&デプロイ監視 (main ブランチ)" `
            | Out-Null

        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  タスクスケジューラへの登録が完了しました！" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "  タスク名: $TaskName" -ForegroundColor Gray
        Write-Host "  トリガー: ログオン時" -ForegroundColor Gray
        Write-Host "  実行内容: AutoBuild.ps1 (30秒間隔で main 監視)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  今すぐ開始するには:" -ForegroundColor Yellow
        Write-Host "    Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
        Write-Host ""
        Write-Host "  削除するには:" -ForegroundColor Yellow
        Write-Host "    .\SetupAutoDeploy.ps1 -Remove" -ForegroundColor White
        Write-Host ""

        # 今すぐ開始するか確認
        $startNow = Read-Host "今すぐ監視を開始しますか？ (Y/N)"
        if ($startNow -eq "Y" -or $startNow -eq "y") {
            Start-ScheduledTask -TaskName $TaskName
            Write-Host ""
            Write-Host "監視を開始しました。" -ForegroundColor Green
        }
    }

    "2" {
        # ========================================
        # 方法2: スタートアップショートカット
        # ========================================

        Write-Host ""
        Write-Host "スタートアップフォルダにショートカットを作成します..." -ForegroundColor Yellow

        $startupFolder = [Environment]::GetFolderPath("Startup")
        $shortcutPath = Join-Path $startupFolder "Tools28_AutoBuild.lnk"

        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "powershell.exe"
        $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$autoBuildScript`""
        $shortcut.WorkingDirectory = $ScriptRoot
        $shortcut.Description = "Tools28 自動ビルド&デプロイ監視"
        $shortcut.Save()

        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  スタートアップへの登録が完了しました！" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ショートカット: $shortcutPath" -ForegroundColor Gray
        Write-Host "  次回ログオン時に監視が自動で開始されます。" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  今すぐ開始するには:" -ForegroundColor Yellow
        Write-Host "    .\AutoBuild.ps1" -ForegroundColor White
        Write-Host ""
        Write-Host "  削除するには:" -ForegroundColor Yellow
        Write-Host "    .\SetupAutoDeploy.ps1 -Remove" -ForegroundColor White
        Write-Host ""
    }

    default {
        Write-Host "無効な選択です。" -ForegroundColor Red
        exit 1
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  自動デプロイ パイプライン:" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. claude/* ブランチを push" -ForegroundColor White
Write-Host "  2. GitHub Actions が main へ自動 squash merge" -ForegroundColor White
Write-Host "  3. AutoBuild.ps1 が main の変更を検知" -ForegroundColor White
Write-Host "  4. 自動で pull → ビルド → Revit にデプロイ" -ForegroundColor White
Write-Host "  5. Windows 通知でビルド結果をお知らせ" -ForegroundColor White
Write-Host ""
Write-Host "  ※ Revit を再起動するとアドインが更新されます" -ForegroundColor Gray
Write-Host ""
