# Tools28 - AutoBuild タスクスケジューラ登録スクリプト
# 使用方法: 管理者権限不要。現在のユーザーのログオン時にAutoBuildをバックグラウンド起動します。
#   .\RegisterAutoBuild.ps1            # タスクを登録
#   .\RegisterAutoBuild.ps1 -Remove    # タスクを削除

param(
    [switch]$Remove
)

$TaskName = "Tools28_AutoBuild"
$ScriptDir = $PSScriptRoot
$VbsPath = Join-Path $ScriptDir "StartAutoBuild.vbs"

if ($Remove) {
    Write-Host "タスク '$TaskName' を削除しています..." -ForegroundColor Yellow
    schtasks /Delete /TN $TaskName /F 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ タスクを削除しました。" -ForegroundColor Green
    } else {
        Write-Host "タスクが見つかりません（既に削除済み）。" -ForegroundColor Gray
    }
    exit
}

# VBSファイルの存在確認
if (-not (Test-Path $VbsPath)) {
    Write-Host "エラー: $VbsPath が見つかりません。" -ForegroundColor Red
    exit 1
}

# 既存タスクの確認
$existingTask = schtasks /Query /TN $TaskName 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "タスク '$TaskName' は既に登録されています。" -ForegroundColor Yellow
    Write-Host "再登録しますか？ (Y/N): " -ForegroundColor Yellow -NoNewline
    $answer = Read-Host
    if ($answer -ne "Y" -and $answer -ne "y") {
        Write-Host "キャンセルしました。" -ForegroundColor Gray
        exit
    }
    schtasks /Delete /TN $TaskName /F 2>$null
}

# タスク登録（ユーザーログオン時に実行）
Write-Host "タスク '$TaskName' を登録しています..." -ForegroundColor Yellow

$taskCommand = "wscript.exe ""$VbsPath"""
schtasks /Create /TN $TaskName /TR $taskCommand /SC ONLOGON /RL LIMITED /F

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "✓ タスクスケジューラに登録しました！" -ForegroundColor Green
    Write-Host ""
    Write-Host "  タスク名: $TaskName" -ForegroundColor Gray
    Write-Host "  トリガー: ログオン時" -ForegroundColor Gray
    Write-Host "  実行: $VbsPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "次回ログオン時から自動的にバックグラウンドで監視が開始されます。" -ForegroundColor Cyan
    Write-Host "今すぐ開始するには: wscript.exe ""$VbsPath""" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "削除する場合: .\RegisterAutoBuild.ps1 -Remove" -ForegroundColor Gray
} else {
    Write-Host ""
    Write-Host "✗ タスク登録に失敗しました。" -ForegroundColor Red
    Write-Host "管理者権限で実行してみてください。" -ForegroundColor Yellow
}
