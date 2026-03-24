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

# ログファイル（バックグラウンド実行時のトラブルシュート用）
$LogFile = Join-Path $PSScriptRoot "AutoBuild.log"

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $Message"
    Write-Host $logLine -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $logLine -ErrorAction SilentlyContinue
}

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

Write-Log "監視を開始します (コミット: $($lastCommit.Substring(0, 7)))" "Green"
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
            $buildCount++
            $commitMsg = git log origin/main -1 --format="%s" 2>$null

            Write-Host ""
            Write-Log "変更を検知！ (ビルド #$buildCount) コミット: $commitMsg" "Yellow"
            Write-Host ""

            # ローカルの未コミット変更を退避してから pull
            Write-Log "pull 開始..." "Yellow"

            # 未コミット変更がある場合は stash で退避
            $stashed = $false
            $statusOutput = git status --porcelain 2>$null
            if ($statusOutput) {
                Write-Log "  未コミット変更を検出 → stash で退避" "Yellow"
                git stash push -u -m "AutoBuild: auto-stash before pull" 2>$null
                $stashed = ($LASTEXITCODE -eq 0)
            }

            git pull origin main 2>$null

            if ($LASTEXITCODE -ne 0) {
                Write-Log "  pull 失敗 → reset --hard + clean で強制同期" "Yellow"
                git reset --hard origin/main 2>$null
                git clean -fd 2>$null
                # reset --hard 後は stash pop 不要（クリーンな状態を維持）
                $stashed = $false
            }

            # stash を戻す（コンフリクトしても無視）
            if ($stashed) {
                Write-Log "  stash を復元中..." "Gray"
                git stash pop 2>$null
                if ($LASTEXITCODE -ne 0) {
                    Write-Log "  stash 復元コンフリクト → stash を破棄 + clean" "Yellow"
                    git checkout -- . 2>$null
                    git clean -fd 2>$null
                    git stash drop 2>$null
                }
            }

            # pull 後の確認
            $localHead = git rev-parse HEAD 2>$null
            if ($localHead -ne $remoteCommit) {
                Write-Log "エラー: pull 後もコミットが一致しません (local=$($localHead.Substring(0,7)) remote=$($remoteCommit.Substring(0,7)))" "Red"

                # 通知（失敗）
                try {
                    Add-Type -AssemblyName System.Windows.Forms
                    $balloon = New-Object System.Windows.Forms.NotifyIcon
                    $balloon.Icon = [System.Drawing.SystemIcons]::Error
                    $balloon.BalloonTipIcon = "Error"
                    $balloon.BalloonTipTitle = "Tools28 ビルド失敗"
                    $balloon.BalloonTipText = "git pull に失敗しました。ローカルリポジトリを確認してください。"
                    $balloon.Visible = $true
                    $balloon.ShowBalloonTip(10000)
                    Start-Sleep -Seconds 3
                    $balloon.Dispose()
                } catch {}

                $lastCommit = $remoteCommit
                Start-Sleep -Seconds $Interval
                continue
            }

            Write-Log "pull 成功 (HEAD: $($localHead.Substring(0,7)))" "Green"

            # ビルド & デプロイ
            Write-Log "ビルド & デプロイ開始..." "Yellow"
            Write-Host ""

            # QuickBuild の出力をログにも記録
            $buildLog = & .\QuickBuild.ps1 2>&1 | Tee-Object -FilePath "$LogFile.build" -Append
            $buildLog | ForEach-Object { Write-Host $_ }
            $buildExitCode = $LASTEXITCODE

            # ビルド失敗時はエラー詳細をログに記録
            if ($buildExitCode -ne 0) {
                Add-Content -Path $LogFile -Value "  --- ビルドエラー詳細 ---" -ErrorAction SilentlyContinue
                $buildLog | ForEach-Object {
                    $line = $_.ToString()
                    if ($line -match "error|エラー|失敗") {
                        Add-Content -Path $LogFile -Value "  $line" -ErrorAction SilentlyContinue
                    }
                }
                Add-Content -Path $LogFile -Value "  --- ビルドエラー詳細 ここまで ---" -ErrorAction SilentlyContinue
            }

            if ($buildExitCode -eq 0) {
                # 通知（成功）
                Write-Log "自動ビルド & デプロイ完了！ Revit を再起動してテストしてください。" "Green"

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
                Write-Log "ビルドに失敗しました。" "Red"

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
