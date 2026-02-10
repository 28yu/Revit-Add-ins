# Tools28 - テスト用デプロイスクリプト
# 使用方法: .\Deploy-For-Testing.ps1 -RevitVersion 2024

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("2022", "2023", "2024", "2025", "2026")]
    [string]$RevitVersion
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tools28 - テスト用デプロイ" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$sourceDll = ".\bin\Release\Revit$RevitVersion\Tools28.dll"
$sourcePdb = ".\bin\Release\Revit$RevitVersion\Tools28.pdb"
$sourceAddin = ".\Addins\Tools28_Revit$RevitVersion.addin"
$targetDir = "$env:APPDATA\Autodesk\Revit\Addins\$RevitVersion\"

# ビルド成果物の確認
if (-not (Test-Path $sourceDll)) {
    Write-Host "エラー: DLLが見つかりません: $sourceDll" -ForegroundColor Red
    Write-Host ""
    Write-Host "以下のコマンドでビルドしてください:" -ForegroundColor Yellow
    Write-Host "  msbuild Tools28.csproj /p:Configuration=Release /p:RevitVersion=$RevitVersion" -ForegroundColor Gray
    Write-Host "または" -ForegroundColor Yellow
    Write-Host "  .\BuildAll.ps1" -ForegroundColor Gray
    exit 1
}

# .addinファイルの確認
if (-not (Test-Path $sourceAddin)) {
    Write-Host "警告: .addinファイルが見つかりません" -ForegroundColor Yellow
    Write-Host "GenerateAddins.ps1 を実行して.addinファイルを生成します..." -ForegroundColor Yellow
    Write-Host ""
    
    & .\GenerateAddins.ps1
    
    if (-not (Test-Path $sourceAddin)) {
        Write-Host "エラー: .addinファイルの生成に失敗しました" -ForegroundColor Red
        exit 1
    }
}

# ターゲットディレクトリの確認
if (-not (Test-Path $targetDir)) {
    Write-Host "エラー: Revit $RevitVersion のアドインフォルダが見つかりません" -ForegroundColor Red
    Write-Host "パス: $targetDir" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Revit $RevitVersion がインストールされているか確認してください。" -ForegroundColor Yellow
    exit 1
}

# バックアップの作成（既存のファイルがある場合）
$targetDll = Join-Path $targetDir "Tools28.dll"
if (Test-Path $targetDll) {
    $backupDir = Join-Path $targetDir "backup"
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupDll = Join-Path $backupDir "Tools28_$timestamp.dll"
    
    Copy-Item $targetDll $backupDll -Force
    Write-Host "既存のDLLをバックアップしました: backup\Tools28_$timestamp.dll" -ForegroundColor Gray
}

# デプロイ実行
Write-Host "デプロイ中..." -ForegroundColor Yellow

try {
    # DLLをコピー
    Copy-Item $sourceDll $targetDir -Force
    Write-Host "✓ Tools28.dll をコピーしました" -ForegroundColor Green
    
    # PDBをコピー（デバッグ用）
    if (Test-Path $sourcePdb) {
        Copy-Item $sourcePdb $targetDir -Force
        Write-Host "✓ Tools28.pdb をコピーしました" -ForegroundColor Green
    }
    
    # .addinファイルをコピー
    $targetAddin = Join-Path $targetDir "Tools28.addin"
    Copy-Item $sourceAddin $targetAddin -Force
    Write-Host "✓ Tools28.addin をコピーしました" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "デプロイが完了しました！" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "次のステップ:" -ForegroundColor Yellow
    Write-Host "1. Revit $RevitVersion を起動してください" -ForegroundColor Gray
    Write-Host "2. リボンに「Tools28」タブが表示されることを確認" -ForegroundColor Gray
    Write-Host "3. 各コマンドをテストしてください" -ForegroundColor Gray
    Write-Host ""
    Write-Host "問題が発生した場合:" -ForegroundColor Yellow
    Write-Host "- Revitのジャーナルファイルを確認" -ForegroundColor Gray
    Write-Host "- Visual Studioのデバッガーをアタッチ" -ForegroundColor Gray
    
} catch {
    Write-Host ""
    Write-Host "エラー: デプロイに失敗しました" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
