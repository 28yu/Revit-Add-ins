# Tools28 diagnostic log uploader for Claude Code
#
# 使い方（ローカルPC / actions-runner の作業ディレクトリで）:
#   cd C:\actions-runner\_work\Revit-Add-ins\Revit-Add-ins
#   git pull origin main
#   .\Send-Tools28Log.ps1
#
# C:\temp\Tools28_debug.txt を .diag\ にコピーして git push し、Claude が読めるようにする。
# （Excelインポート等の診断ログ。DiagLog.Write の出力先）

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$srcDir   = "C:\temp"
$dstDir   = Join-Path $repoRoot ".diag"

$src = Join-Path $srcDir "Tools28_debug.txt"
if (-not (Test-Path $src)) {
    Write-Host "ERROR: $src が見つかりません" -ForegroundColor Red
    Write-Host "  先に Revit で対象コマンド（Excelインポート等）を実行してください。" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $dstDir)) {
    New-Item -ItemType Directory -Path $dstDir | Out-Null
}

$dst = Join-Path $dstDir "Tools28_debug.txt"
Copy-Item $src $dst -Force
$sizeKB = [math]::Round((Get-Item $dst).Length / 1024, 1)
Write-Host ("  Copied: Tools28_debug.txt ({0} KB)" -f $sizeKB) -ForegroundColor Cyan

# --- git add / commit / push ---
Push-Location $repoRoot
try {
    git add -f ".diag\Tools28_debug.txt"
    git commit -m "diag: Tools28_debug ログを更新"
    git push
    Write-Host "Push 完了。Claude に『ログを更新した』と伝えてください。" -ForegroundColor Green
}
catch {
    Write-Host ("Git error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
