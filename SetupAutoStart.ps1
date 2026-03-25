# Tools28 - AutoBuild 自動起動セットアップ
# 使用方法: 管理者権限不要。一度だけ実行すればOK。
#
# 実行すると:
#   Windows ログオン時に AutoBuild.ps1 が自動起動するようになります。
#   スタートアップフォルダにショートカットを作成します。
#
# 解除: .\SetupAutoStart.ps1 -Remove

param(
    [switch]$Remove
)

$startupFolder = [Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startupFolder "Tools28 AutoBuild.lnk"

if ($Remove) {
    if (Test-Path $shortcutPath) {
        Remove-Item $shortcutPath -Force
        Write-Host "自動起動を解除しました。" -ForegroundColor Green
        Write-Host "  削除: $shortcutPath" -ForegroundColor Gray
    } else {
        Write-Host "自動起動は設定されていません。" -ForegroundColor Yellow
    }
    return
}

# リポジトリルートの確認
$repoRoot = $PSScriptRoot
$autoBuildScript = Join-Path $repoRoot "AutoBuild.ps1"

if (-not (Test-Path $autoBuildScript)) {
    Write-Host "エラー: AutoBuild.ps1 が見つかりません: $autoBuildScript" -ForegroundColor Red
    exit 1
}

# ショートカットの作成
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$autoBuildScript`""
$shortcut.WorkingDirectory = $repoRoot
$shortcut.Description = "Tools28 AutoBuild - main ブランチ監視 & 自動ビルド"
$shortcut.WindowStyle = 1  # 通常ウィンドウ
$shortcut.Save()

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  AutoBuild 自動起動を設定しました" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  ショートカット: $shortcutPath" -ForegroundColor Gray
Write-Host "  対象スクリプト: $autoBuildScript" -ForegroundColor Gray
Write-Host ""
Write-Host "  次回の Windows ログオンから AutoBuild が自動起動します。" -ForegroundColor Green
Write-Host ""
Write-Host "  解除するには:" -ForegroundColor Yellow
Write-Host "    .\SetupAutoStart.ps1 -Remove" -ForegroundColor White
Write-Host ""
