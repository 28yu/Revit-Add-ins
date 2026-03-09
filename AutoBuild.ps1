# Tools28 - 自動ビルド & デプロイ監視スクリプト
# 使用方法:
#   .\AutoBuild.ps1                     # デフォルト設定で起動（30秒間隔）
#   .\AutoBuild.ps1 -Interval 60        # 60秒間隔で監視
#
# 概要:
#   main ブランチのリモート更新を監視し、変更があれば自動で pull → ビルド → デプロイします。
#   Claude Code で開発指示 → push → auto-merge → main 更新 → このスクリプトが検知 → 自動ビルド
#
# 終了: Ctrl+C

param(
    [Parameter(Mandatory=$false)]
    [int]$Interval = 30
)

# ========================================
# 初期設定
# ========================================

$host.UI.RawUI.WindowTitle = "Tools28 AutoBuild - 監視中"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Tools28 - 自動ビルド & デプロイ監視" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  監視間隔: ${Interval}秒" -ForegroundColor Gray
Write-Host "  終了: Ctrl+C" -ForegroundColor Gray
Write-Host ""

# リポジトリのルートにいることを確認
if (-not (Test-Path ".\Tools28.csproj")) {
    Write-Host "エラー: Tools28.csproj が見つかりません。" -ForegroundColor Red
    Write-Host "リポジトリのルートディレクトリで実行してください。" -ForegroundColor Yellow
    exit 1
}

# 現在のブランチを main に切り替え（必要な場合）
$currentBranch = git rev-parse --abbrev-ref HEAD 2>$null
if ($currentBranch -ne "main") {
    Write-Host "ブランチを main に切り替えます..." -ForegroundColor Yellow
    git checkout main 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "エラー: main ブランチへの切り替えに失敗しました。" -ForegroundColor Red
        exit 1
    }
}

# 初回 fetch
git fetch origin main 2>$null
$lastCommit = git rev-parse origin/main 2>$null

Write-Host "監視を開始します..." -ForegroundColor Green
Write-Host "  現在のコミット: $($lastCommit.Substring(0, 7))" -ForegroundColor Gray
Write-Host ""

# ========================================
# 監視ループ
# ========================================

$buildCount = 0

while ($true) {
    try {
        # リモートの最新状態を取得
        git fetch origin main 2>$null

        if ($LASTEXITCODE -ne 0) {
            $timestamp = Get-Date -Format "HH:mm:ss"
            Write-Host "[$timestamp] fetch 失敗（ネットワーク確認）" -ForegroundColor Yellow
            Start-Sleep -Seconds $Interval
            continue
        }

        $remoteCommit = git rev-parse origin/main 2>$null

        # 変更があるか確認
        if ($remoteCommit -ne $lastCommit) {
            $timestamp = Get-Date -Format "HH:mm:ss"
            $buildCount++

            Write-Host ""
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "[$timestamp] 変更を検知！ (ビルド #$buildCount)" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow

            # コミットメッセージを表示
            $commitMsg = git log origin/main -1 --format="%s" 2>$null
            Write-Host "  コミット: $commitMsg" -ForegroundColor Gray
            Write-Host ""

            # pull
            Write-Host "pull 中..." -ForegroundColor Yellow
            git pull origin main 2>$null

            if ($LASTEXITCODE -ne 0) {
                Write-Host "エラー: pull に失敗しました。" -ForegroundColor Red

                # 通知（失敗）
                try {
                    Add-Type -AssemblyName System.Windows.Forms
                    $balloon = New-Object System.Windows.Forms.NotifyIcon
                    $balloon.Icon = [System.Drawing.SystemIcons]::Error
                    $balloon.BalloonTipIcon = "Error"
                    $balloon.BalloonTipTitle = "Tools28 ビルド失敗"
                    $balloon.BalloonTipText = "git pull に失敗しました。"
                    $balloon.Visible = $true
                    $balloon.ShowBalloonTip(10000)
                    Start-Sleep -Seconds 3
                    $balloon.Dispose()
                } catch {}

                $lastCommit = $remoteCommit
                Start-Sleep -Seconds $Interval
                continue
            }

            # ビルド & デプロイ
            Write-Host "ビルド & デプロイ中..." -ForegroundColor Yellow
            Write-Host ""

            & .\QuickBuild.ps1

            if ($LASTEXITCODE -eq 0) {
                # 通知（成功）
                Write-Host ""
                Write-Host "========================================" -ForegroundColor Green
                Write-Host "  自動ビルド & デプロイ完了！" -ForegroundColor Green
                Write-Host "  Revit を再起動してテストしてください。" -ForegroundColor Green
                Write-Host "========================================" -ForegroundColor Green

                try {
                    Add-Type -AssemblyName System.Windows.Forms
                    $balloon = New-Object System.Windows.Forms.NotifyIcon
                    $balloon.Icon = [System.Drawing.SystemIcons]::Information
                    $balloon.BalloonTipIcon = "Info"
                    $balloon.BalloonTipTitle = "Tools28 ビルド完了"
                    $balloon.BalloonTipText = "ビルド & デプロイが完了しました。Revit を再起動してテストしてください。"
                    $balloon.Visible = $true
                    $balloon.ShowBalloonTip(10000)
                    Start-Sleep -Seconds 3
                    $balloon.Dispose()
                } catch {}
            } else {
                # 通知（失敗）
                Write-Host ""
                Write-Host "========================================" -ForegroundColor Red
                Write-Host "  ビルドに失敗しました。" -ForegroundColor Red
                Write-Host "========================================" -ForegroundColor Red

                try {
                    Add-Type -AssemblyName System.Windows.Forms
                    $balloon = New-Object System.Windows.Forms.NotifyIcon
                    $balloon.Icon = [System.Drawing.SystemIcons]::Error
                    $balloon.BalloonTipIcon = "Error"
                    $balloon.BalloonTipTitle = "Tools28 ビルド失敗"
                    $balloon.BalloonTipText = "ビルドに失敗しました。ログを確認してください。"
                    $balloon.Visible = $true
                    $balloon.ShowBalloonTip(10000)
                    Start-Sleep -Seconds 3
                    $balloon.Dispose()
                } catch {}
            }

            $lastCommit = $remoteCommit
            Write-Host ""
            Write-Host "監視を再開します..." -ForegroundColor Gray
        } else {
            # 変更なし - ドットで監視中を表示
            Write-Host "." -NoNewline -ForegroundColor DarkGray
        }

        Start-Sleep -Seconds $Interval
    }
    catch {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host ""
        Write-Host "[$timestamp] エラー: $_" -ForegroundColor Red
        Start-Sleep -Seconds $Interval
    }
}
