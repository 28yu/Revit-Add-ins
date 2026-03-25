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

            # ローカルを origin/main に強制同期（ローカル変更は破棄）
            Write-Log "pull 開始..." "Yellow"

            git reset --hard origin/main 2>$null
            git clean -fd 2>$null

            # pull 後の確認
            $localHead = git rev-parse HEAD 2>$null
            if ($localHead -ne $remoteCommit) {
                Write-Log "エラー: pull 後もコミットが一致しません (local=$($localHead.Substring(0,7)) remote=$($remoteCommit.Substring(0,7)))" "Red"

                # 通知（失敗）— 別プロセスで表示（監視ループをブロックしない）
                Start-Process powershell -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-Command', "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('git pull に失敗しました。`nローカルリポジトリを確認してください。', 'Tools28 ビルド失敗', 'OK', 'Error')" -WindowStyle Hidden

                $lastCommit = $remoteCommit
                Start-Sleep -Seconds $Interval
                continue
            }

            Write-Log "pull 成功 (HEAD: $($localHead.Substring(0,7)))" "Green"

            # ビルド & デプロイ
            Write-Log "ビルド & デプロイ開始..." "Yellow"
            Write-Host ""

            # QuickBuild の出力をキャプチャ
            $buildLog = & .\QuickBuild.ps1 *>&1
            $buildLog | ForEach-Object { Write-Host $_ }
            $buildLog | Out-File -FilePath "$LogFile.build" -Append -Encoding UTF8

            # ビルド成功判定: $LASTEXITCODE は不安定なため、出力内容で判定する
            $buildOutput = ($buildLog | ForEach-Object { $_.ToString() }) -join "`n"
            $buildSuccess = $buildOutput -match "完了しました"

            # ビルド失敗時は全出力をログに記録（原因特定用）
            if (-not $buildSuccess) {
                Add-Content -Path $LogFile -Value "  --- ビルド全出力 ---" -ErrorAction SilentlyContinue
                $buildLog | ForEach-Object {
                    $line = $_.ToString()
                    Add-Content -Path $LogFile -Value "  $line" -ErrorAction SilentlyContinue
                }
                Add-Content -Path $LogFile -Value "  --- ビルド全出力 ここまで ---" -ErrorAction SilentlyContinue
            }

            if ($buildSuccess) {
                # 通知（成功）
                Write-Log "自動ビルド & デプロイ完了！ Revit を再起動してテストしてください。" "Green"

                # 通知（成功）— 別プロセスで表示（監視ループをブロックしない）
                Start-Process powershell -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-Command', "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('ビルド & デプロイが完了しました。`nRevit を再起動してテストしてください。', 'Tools28 ビルド完了', 'OK', 'Information')" -WindowStyle Hidden
            } else {
                # 通知（失敗）
                Write-Log "ビルドに失敗しました。" "Red"

                # 通知（失敗）— 別プロセスで表示（監視ループをブロックしない）
                Start-Process powershell -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-Command', "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('ビルドに失敗しました。`nログを確認してください。', 'Tools28 ビルド失敗', 'OK', 'Error')" -WindowStyle Hidden
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
