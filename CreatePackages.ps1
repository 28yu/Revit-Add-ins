# Tools28 - 配布インストーラ作成スクリプト
# 使用方法: .\CreatePackages.ps1 -Version "2.6"
# 前提: .\BuildAll.ps1 を事前に実行して全6バージョンの DLL をビルド済みであること
#
# 出力: .\Dist\28Tools_Setup_vX.X.exe（Revit 2021〜2026 を 1 本にまとめた Inno Setup インストーラ）
#   インストール先: C:\ProgramData\Autodesk\Revit\Addins\<ver>\Tools28.addin
#                   C:\ProgramData\Autodesk\Revit\Addins\<ver>\28Tools\*.dll
#   定義は Installer\28Tools.iss。リリース物は CI（build-and-release.yml の installer ジョブ）が作る。

param(
    [string]$Version = "0.0"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tools28 - 配布インストーラ作成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "バージョン: v$Version" -ForegroundColor Cyan
Write-Host ""

$revitVersions = @("2021", "2022", "2023", "2024", "2025", "2026")
$issPath = ".\Installer\28Tools.iss"
$outputDir = ".\Dist"

# ビルド成果物の確認（1つでも欠けていると ISCC が失敗する）
$missing = @()
foreach ($ver in $revitVersions) {
    $dll = ".\bin\Release\Revit$ver\Tools28.dll"
    if (Test-Path $dll) {
        Write-Host "  ✓ Revit $ver : $dll" -ForegroundColor Green
    } else {
        Write-Host "  ✗ Revit $ver : DLL が見つかりません ($dll)" -ForegroundColor Red
        $missing += $ver
    }
}
if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "先に .\BuildAll.ps1 を実行してください" -ForegroundColor Yellow
    exit 1
}

# Inno Setup のコンパイラ（ISCC.exe）を探す
$iscc = (Get-Command iscc -ErrorAction SilentlyContinue).Source
if (-not $iscc) {
    foreach ($p in @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe")) {
        if (Test-Path $p) { $iscc = $p; break }
    }
}
if (-not $iscc) {
    Write-Host ""
    Write-Host "Inno Setup 6 が見つかりません。次のどちらかでインストールしてから再実行してください:" -ForegroundColor Yellow
    Write-Host "  - https://jrsoftware.org/isdl.php からダウンロード（6.5 以降を推奨：中国語の言語ファイルが同梱）" -ForegroundColor Gray
    Write-Host "  - winget install JRSoftware.InnoSetup" -ForegroundColor Gray
    Write-Host ""
    Write-Host "手動でビルドする場合: ISCC.exe /DAppVersion=$Version $issPath" -ForegroundColor Gray
    exit 1
}
Write-Host ""
Write-Host "ISCC: $iscc" -ForegroundColor Green

# コンパイル（出力は Installer\Output\）
& $iscc "/DAppVersion=$Version" $issPath
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ インストーラのビルドに失敗しました ($LASTEXITCODE)" -ForegroundColor Red
    exit 1
}

$exeName = "28Tools_Setup_v$Version.exe"
$built = ".\Installer\Output\$exeName"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}
Copy-Item -Path $built -Destination $outputDir -Force

$size = [math]::Round((Get-Item (Join-Path $outputDir $exeName)).Length / 1MB, 2)
Write-Host ""
Write-Host "✓ $exeName を作成しました ($size MB)" -ForegroundColor Green
Write-Host "出力先: $outputDir" -ForegroundColor Cyan
Write-Host ""
