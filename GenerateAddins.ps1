# Tools28 - .addinファイル生成スクリプト
# 使用方法: .\GenerateAddins.ps1

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tools28 - .addinファイル生成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$revitVersions = @("2022", "2023", "2024", "2025", "2026")
$outputDir = ".\Addins"

# 出力ディレクトリを作成
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
    Write-Host "出力ディレクトリを作成しました: $outputDir" -ForegroundColor Green
}

foreach ($version in $revitVersions) {
    $addinContent = @"
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>Tools28</Name>
    <Assembly>Tools28.dll</Assembly>
    <AddInId>E40A22D9-DD80-4387-BA5D-8F9E4DF157D6</AddInId>
    <FullClassName>Tools28.Application</FullClassName>
    <VendorId>Tools28</VendorId>
    <VendorDescription>Tools28, https://tools28.com</VendorDescription>
  </AddIn>
</RevitAddIns>
"@

    $fileName = Join-Path $outputDir "Tools28_Revit$version.addin"
    $addinContent | Out-File -FilePath $fileName -Encoding UTF8
    Write-Host "✓ 生成完了: Tools28_Revit$version.addin" -ForegroundColor Green
}

Write-Host ""
Write-Host "すべての.addinファイルが生成されました: $outputDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "次のステップ:" -ForegroundColor Yellow
Write-Host "1. BuildAll.ps1 を実行してすべてのバージョンをビルド" -ForegroundColor Gray
Write-Host "2. 各バージョンのDLLを対応する.addinファイルと一緒に配布" -ForegroundColor Gray
