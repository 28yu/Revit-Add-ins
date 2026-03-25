# Tools28 - クイックビルド＆デプロイスクリプト
# 使用方法:
#   .\QuickBuild.ps1                    # dev-config.jsonの設定を使用
#   .\QuickBuild.ps1 -RevitVersion 2024 # 特定バージョンを指定

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("2021", "2022", "2023", "2024", "2025", "2026")]
    [string]$RevitVersion
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Tools28 - QuickBuild & Deploy" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Revitバージョンの決定
if (-not $RevitVersion) {
    # dev-config.jsonから読み込み
    $configPath = ".\dev-config.json"
    if (Test-Path $configPath) {
        try {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            $RevitVersion = $config.defaultRevitVersion
            Write-Host "設定ファイルから読み込み: Revit $RevitVersion" -ForegroundColor Gray
        } catch {
            Write-Host "警告: dev-config.jsonの読み込みに失敗しました" -ForegroundColor Yellow
            $RevitVersion = "2022"
            Write-Host "デフォルト値を使用: Revit $RevitVersion" -ForegroundColor Gray
        }
    } else {
        $RevitVersion = "2022"
        Write-Host "デフォルト値を使用: Revit $RevitVersion" -ForegroundColor Gray
    }
} else {
    Write-Host "指定されたバージョン: Revit $RevitVersion" -ForegroundColor Gray
}

Write-Host ""

# ========================================
# ステップ1: ビルド
# ========================================

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ステップ1: ビルド" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# MSBuildのパスを検索
$msbuildPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" `
    -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe `
    -prerelease | Select-Object -First 1

if (-not $msbuildPath) {
    Write-Host "エラー: MSBuildが見つかりません" -ForegroundColor Red
    Write-Host "Visual Studio 2017以降がインストールされている必要があります" -ForegroundColor Yellow
    exit 1
}

Write-Host "MSBuild: $msbuildPath" -ForegroundColor Gray
Write-Host "ビルド中..." -ForegroundColor Yellow
Write-Host ""

$env:RevitVersion = $RevitVersion

try {
    & $msbuildPath "Tools28.csproj" `
        /p:Configuration=Release `
        /p:RevitVersion=$RevitVersion `
        /p:Platform=AnyCPU `
        /v:minimal `
        /nologo `
        /restore

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "✗ ビルドに失敗しました" -ForegroundColor Red
        exit 1
    }

    Write-Host ""
    Write-Host "✓ ビルドに成功しました" -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "✗ ビルドに失敗しました: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""

# ========================================
# ステップ2: デプロイ
# ========================================

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ステップ2: デプロイ" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$sourceDll = ".\bin\Release\Revit$RevitVersion\Tools28.dll"
$sourcePdb = ".\bin\Release\Revit$RevitVersion\Tools28.pdb"
$sourceAddin = ".\Packages\$RevitVersion\28Tools\Tools28.addin"
$targetDir = "$env:ProgramData\Autodesk\Revit\Addins\$RevitVersion\"
$targetToolsDir = "$env:ProgramData\Autodesk\Revit\Addins\$RevitVersion\28Tools\"

# ビルド成果物の確認
if (-not (Test-Path $sourceDll)) {
    Write-Host "エラー: DLLが見つかりません: $sourceDll" -ForegroundColor Red
    exit 1
}

# .addinファイルの確認
if (-not (Test-Path $sourceAddin)) {
    Write-Host "エラー: .addinファイルが見つかりません: $sourceAddin" -ForegroundColor Red
    exit 1
}

# ターゲットディレクトリの作成（存在しない場合）
if (-not (Test-Path $targetDir)) {
    Write-Host "ターゲットディレクトリを作成: $targetDir" -ForegroundColor Gray
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

# 28Toolsサブフォルダの作成
if (-not (Test-Path $targetToolsDir)) {
    New-Item -ItemType Directory -Path $targetToolsDir -Force | Out-Null
}

# デプロイ実行
Write-Host "デプロイ中..." -ForegroundColor Yellow
Write-Host ""

try {
    # 全DLLを 28Tools サブフォルダにコピー（Tools28.dll + ClosedXML等の依存ライブラリ）
    $buildOutputDir = ".\bin\Release\Revit$RevitVersion\"
    $allDlls = Get-ChildItem -Path $buildOutputDir -Filter "*.dll"
    $copyFailed = @()
    foreach ($dll in $allDlls) {
        try {
            Copy-Item $dll.FullName $targetToolsDir -Force -ErrorAction Stop
            Write-Host "✓ $($dll.Name) をコピーしました" -ForegroundColor Green
        } catch {
            $copyFailed += $dll.Name
            Write-Host "⚠ $($dll.Name) はロックされているためスキップしました（既存ファイルを使用）" -ForegroundColor Yellow
        }
    }

    # PDBをコピー（デバッグ用、ロック時はスキップ）
    if (Test-Path $sourcePdb) {
        try {
            Copy-Item $sourcePdb $targetToolsDir -Force -ErrorAction Stop
            Write-Host "✓ Tools28.pdb をコピーしました" -ForegroundColor Green
        } catch {
            Write-Host "⚠ Tools28.pdb はロックされているためスキップしました" -ForegroundColor Yellow
        }
    }

    # .addinファイルをAddinsルートにコピー
    $targetAddin = Join-Path $targetDir "Tools28.addin"
    Copy-Item $sourceAddin $targetAddin -Force
    Write-Host "✓ Tools28.addin をコピーしました" -ForegroundColor Green

    # 旧バージョンのクリーンアップ（ルートに直置きされていた場合）
    $oldRootDll = Join-Path $targetDir "Tools28.dll"
    if (Test-Path $oldRootDll) {
        Remove-Item $oldRootDll -Force
        Write-Host "✓ 旧 Tools28.dll（ルート直置き）を削除しました" -ForegroundColor Yellow
    }

    # Tools28.dll のコピーに失敗した場合は致命的エラー
    if ($copyFailed -contains "Tools28.dll") {
        throw "Tools28.dll のコピーに失敗しました。Revit を閉じてから再実行してください。"
    }

    if ($copyFailed.Count -gt 0) {
        Write-Host ""
        Write-Host "⚠ 一部のDLL ($($copyFailed -join ', ')) はロック中のためスキップされました" -ForegroundColor Yellow
        Write-Host "  Revit 再起動後に再度実行すると全ファイルが更新されます" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "完了しました！" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "デプロイ先: $targetDir" -ForegroundColor Gray
    Write-Host ""
    Write-Host "次のステップ:" -ForegroundColor Yellow
    Write-Host "  1. Revit $RevitVersion を起動（または再起動）してください" -ForegroundColor White
    Write-Host "  2. リボンに「28 Tools」タブが表示されることを確認" -ForegroundColor White
    Write-Host "  3. 機能をテストしてください" -ForegroundColor White
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "✗ デプロイに失敗しました" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

exit 0
