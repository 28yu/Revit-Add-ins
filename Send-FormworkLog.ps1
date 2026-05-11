# Formwork 診断ログを Claude Code に共有するためのスクリプト
#
# 使い方:
#   1. Revit で「型枠数量算出」を実行 → C:\temp\Formwork_debug.txt が更新される
#   2. このスクリプトを実行: .\Send-FormworkLog.ps1
#   3. Claude に "ログを更新したので確認してください" と伝える
#
# 動作:
#   - C:\temp\Formwork_debug.txt を .diag\Formwork_debug.txt にコピー
#   - 現在のブランチに commit して push
#   - Claude が git pull すれば直接 Read ツールでログを読める

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$srcLog   = "C:\temp\Formwork_debug.txt"
$dstDir   = Join-Path $repoRoot ".diag"
$dstLog   = Join-Path $dstDir   "Formwork_debug.txt"

if (-not (Test-Path $srcLog)) {
    Write-Host "エラー: $srcLog が見つかりません。Revit で型枠数量算出を先に実行してください。" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $dstDir)) {
    New-Item -ItemType Directory -Path $dstDir | Out-Null
}

Copy-Item $srcLog $dstLog -Force

$size = (Get-Item $dstLog).Length
$sizeKB = [math]::Round($size / 1KB, 1)
Write-Host "ログをコピー: $dstLog ($sizeKB KB)" -ForegroundColor Green

Push-Location $repoRoot
try {
    git add -f .\.diag\Formwork_debug.txt
    git commit -m "診断ログ更新: Formwork_debug.txt"
    git push
    Write-Host "push 完了。Claude に「ログを更新しました」と伝えてください。" -ForegroundColor Green
}
finally {
    Pop-Location
}
