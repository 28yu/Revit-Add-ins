# Tools28 - 配布パッケージ作成スクリプト
# 使用方法: .\CreatePackages.ps1
# 前提: .\BuildAll.ps1 を事前に実行してDLLをビルド済みであること
#
# 出力される配布ZIP構成:
#   28Tools_Revit20XX_vX.X.zip
#   ├── 28Tools/
#   │   ├── Tools28.dll
#   │   └── Tools28.addin
#   ├── install.bat
#   ├── uninstall.bat
#   └── README.txt

param(
    [string]$Version = "1.0"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tools28 - 配布パッケージ作成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "バージョン: v$Version" -ForegroundColor Cyan
Write-Host ""

$revitVersions = @("2021", "2022", "2023", "2024", "2025", "2026")
$packageTemplateDir = ".\Packages"
$outputDir = ".\Dist"

# 出力ディレクトリを作成
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# 各バージョンのパッケージを作成
$successCount = 0
$failCount = 0

foreach ($ver in $revitVersions) {
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    Write-Host "Revit $ver 用パッケージを作成中..." -ForegroundColor Yellow
    Write-Host "----------------------------------------" -ForegroundColor Yellow

    $buildDll = ".\bin\Release\Revit$ver\Tools28.dll"
    $templateDir = Join-Path $packageTemplateDir $ver

    # ビルド成果物の確認
    if (-not (Test-Path $buildDll)) {
        Write-Host "  ✗ DLLが見つかりません: $buildDll" -ForegroundColor Red
        Write-Host "  先に .\BuildAll.ps1 を実行してください" -ForegroundColor Yellow
        $failCount++
        continue
    }

    # テンプレートディレクトリの確認
    if (-not (Test-Path $templateDir)) {
        Write-Host "  ✗ テンプレートが見つかりません: $templateDir" -ForegroundColor Red
        $failCount++
        continue
    }

    # 一時ディレクトリで作業
    $zipName = "28Tools_Revit${ver}_v${Version}"
    $tempDir = Join-Path $outputDir "_temp_$ver"

    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    # 28Tools/ サブフォルダを作成
    $tools28Dir = Join-Path $tempDir "28Tools"
    New-Item -ItemType Directory -Path $tools28Dir | Out-Null

    # DLL を 28Tools/ にコピー
    Copy-Item -Path $buildDll -Destination (Join-Path $tools28Dir "Tools28.dll")
    Write-Host "  ✓ Tools28.dll をコピーしました" -ForegroundColor Green

    # .addin を 28Tools/ にコピー
    $addinSource = Join-Path $templateDir "28Tools\Tools28.addin"
    if (Test-Path $addinSource) {
        Copy-Item -Path $addinSource -Destination (Join-Path $tools28Dir "Tools28.addin")
        Write-Host "  ✓ Tools28.addin をコピーしました" -ForegroundColor Green
    } else {
        Write-Host "  ✗ .addinファイルが見つかりません: $addinSource" -ForegroundColor Red
        $failCount++
        Remove-Item -Path $tempDir -Recurse -Force
        continue
    }

    # install.bat, uninstall.bat, README.txt をルートにコピー
    foreach ($file in @("install.bat", "uninstall.bat", "README.txt")) {
        $src = Join-Path $templateDir $file
        if (Test-Path $src) {
            Copy-Item -Path $src -Destination (Join-Path $tempDir $file)
        } else {
            Write-Host "  ⚠ $file が見つかりません: $src" -ForegroundColor Yellow
        }
    }

    # ZIP 作成
    $zipPath = Join-Path $outputDir "$zipName.zip"
    if (Test-Path $zipPath) {
        Remove-Item -Path $zipPath -Force
    }

    Compress-Archive -Path "$tempDir\*" -DestinationPath $zipPath -CompressionLevel Optimal
    Write-Host "  ✓ $zipName.zip を作成しました" -ForegroundColor Green

    # 一時ディレクトリを削除
    Remove-Item -Path $tempDir -Recurse -Force

    $successCount++
    Write-Host ""
}

# 結果サマリー
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "パッケージ作成結果" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "成功: $successCount / 失敗: $failCount" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Yellow" })
Write-Host ""

if ($successCount -gt 0) {
    Write-Host "作成されたパッケージ:" -ForegroundColor Yellow
    Get-ChildItem -Path $outputDir -Filter "*.zip" | ForEach-Object {
        $size = [math]::Round($_.Length / 1KB, 2)
        Write-Host "  $($_.Name) ($size KB)" -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "出力先: $outputDir" -ForegroundColor Cyan
    Write-Host "これらのZIPファイルを配布サイトにアップロードしてください。" -ForegroundColor Cyan
}

Write-Host ""
