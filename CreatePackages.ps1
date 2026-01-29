# Tools28 - 配布パッケージ作成スクリプト
# 使用方法: .\CreatePackages.ps1

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tools28 - 配布パッケージ作成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$revitVersions = @("2022", "2023", "2024", "2025", "2026")
$packageDir = ".\Packages"
$addinDir = ".\Addins"

# パッケージディレクトリを作成
if (-not (Test-Path $packageDir)) {
    New-Item -ItemType Directory -Path $packageDir | Out-Null
    Write-Host "パッケージディレクトリを作成しました: $packageDir" -ForegroundColor Green
    Write-Host ""
}

# 各バージョンのパッケージを作成
foreach ($version in $revitVersions) {
    Write-Host "Revit $version 用パッケージを作成中..." -ForegroundColor Yellow
    
    $buildPath = ".\bin\Release\Revit$version"
    $addinFile = Join-Path $addinDir "Tools28_Revit$version.addin"
    $packageName = "Tools28_Revit$version"
    $packagePath = Join-Path $packageDir $packageName
    
    # ビルド成果物が存在するか確認
    if (-not (Test-Path $buildPath)) {
        Write-Host "  ✗ ビルド成果物が見つかりません: $buildPath" -ForegroundColor Red
        Write-Host "  先に BuildAll.ps1 を実行してください" -ForegroundColor Yellow
        continue
    }
    
    # .addinファイルが存在するか確認
    if (-not (Test-Path $addinFile)) {
        Write-Host "  ✗ .addinファイルが見つかりません: $addinFile" -ForegroundColor Red
        Write-Host "  先に GenerateAddins.ps1 を実行してください" -ForegroundColor Yellow
        continue
    }
    
    # パッケージディレクトリを作成
    if (Test-Path $packagePath) {
        Remove-Item -Path $packagePath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $packagePath | Out-Null
    
    # DLLとPDBをコピー
    Copy-Item -Path (Join-Path $buildPath "Tools28.dll") -Destination $packagePath
    $pdbPath = Join-Path $buildPath "Tools28.pdb"
    if (Test-Path $pdbPath) {
        Copy-Item -Path $pdbPath -Destination $packagePath
    }
    
    # .addinファイルをコピー
    Copy-Item -Path $addinFile -Destination (Join-Path $packagePath "Tools28.addin")
    
    # READMEを作成
    $readmeContent = @"
Tools28 - Revit $version 版
====================================

インストール方法
----------------
1. Tools28.dll と Tools28.addin を以下のフォルダにコピーしてください：
   %APPDATA%\Autodesk\Revit\Addins\$version\

2. Revit $version を再起動してください

3. リボンに「Tools28」タブが表示されます

機能
----
- シート作成
- ビューコピー
- グリッドバブル
- セクションボックスコピー
- ビューポート位置調整
- クロップボックスコピー

サポート
--------
問題が発生した場合は、GitHubのIssuesまでお願いします。
https://github.com/[your-username]/Tools28/issues
"@
    
    $readmeContent | Out-File -FilePath (Join-Path $packagePath "README.txt") -Encoding UTF8
    
    # ZIPファイルを作成
    $zipPath = Join-Path $packageDir "$packageName.zip"
    if (Test-Path $zipPath) {
        Remove-Item -Path $zipPath -Force
    }
    
    Compress-Archive -Path "$packagePath\*" -DestinationPath $zipPath -CompressionLevel Optimal
    
    Write-Host "  ✓ パッケージ作成完了: $packageName.zip" -ForegroundColor Green
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "すべてのパッケージが作成されました！" -ForegroundColor Cyan
Write-Host "場所: $packageDir" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# パッケージ一覧を表示
Write-Host "作成されたパッケージ:" -ForegroundColor Yellow
Get-ChildItem -Path $packageDir -Filter "*.zip" | ForEach-Object {
    $size = [math]::Round($_.Length / 1KB, 2)
    Write-Host "  $($_.Name) ($size KB)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "これらのZIPファイルを配布サイトにアップロードできます。" -ForegroundColor Cyan
