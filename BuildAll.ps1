# Tools28 - 全バージョンビルドスクリプト
# 使用方法: .\BuildAll.ps1

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tools28 - Multi-Version Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ビルドするRevitバージョンのリスト
$revitVersions = @("2021", "2022", "2023", "2024", "2025", "2026")

# MSBuildのパスを検索
$msbuildPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" `
    -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe `
    -prerelease | Select-Object -First 1

if (-not $msbuildPath) {
    Write-Host "エラー: MSBuildが見つかりません" -ForegroundColor Red
    Write-Host "Visual Studio 2017以降がインストールされている必要があります" -ForegroundColor Yellow
    exit 1
}

Write-Host "MSBuild: $msbuildPath" -ForegroundColor Green
Write-Host ""

# 各バージョンをビルド
$successCount = 0
$failCount = 0
$results = @()

foreach ($version in $revitVersions) {
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    Write-Host "Revit $version をビルド中..." -ForegroundColor Yellow
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    
    $env:RevitVersion = $version
    
    try {
        & $msbuildPath "Tools28.csproj" `
            /p:Configuration=Release `
            /p:RevitVersion=$version `
            /p:Platform=AnyCPU `
            /v:minimal `
            /nologo
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Revit $version のビルドに成功しました" -ForegroundColor Green
            $successCount++
            $results += @{Version = $version; Status = "成功"; Color = "Green"}
        } else {
            Write-Host "✗ Revit $version のビルドに失敗しました" -ForegroundColor Red
            $failCount++
            $results += @{Version = $version; Status = "失敗"; Color = "Red"}
        }
    } catch {
        Write-Host "✗ Revit $version のビルドに失敗しました: $_" -ForegroundColor Red
        $failCount++
        $results += @{Version = $version; Status = "失敗"; Color = "Red"}
    }
    
    Write-Host ""
}

# 結果サマリー
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ビルド結果" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

foreach ($result in $results) {
    $status = $result.Status
    $color = $result.Color
    Write-Host "Revit $($result.Version): $status" -ForegroundColor $color
}

Write-Host ""
Write-Host "成功: $successCount / 失敗: $failCount" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Yellow" })

# 出力フォルダの一覧を表示
if ($successCount -gt 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "ビルド成果物の場所" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    foreach ($result in $results) {
        if ($result.Status -eq "成功") {
            $outputPath = "bin\Release\Revit$($result.Version)"
            Write-Host "Revit $($result.Version): $outputPath" -ForegroundColor Gray
        }
    }
}

Write-Host ""
Write-Host "完了しました。" -ForegroundColor Cyan

exit $(if ($failCount -eq 0) { 0 } else { 1 })
