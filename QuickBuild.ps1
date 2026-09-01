# Tools28 - クイックビルド＆デプロイスクリプト
# 使用方法:
#   .\QuickBuild.ps1                    # dev-config.jsonの設定を使用
#   .\QuickBuild.ps1 -RevitVersion 2024 # 特定バージョンを指定

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("2021", "2022", "2023", "2024", "2025", "2026")]
    [string]$RevitVersion
)

# デプロイ共通処理（Revit 起動中のロック対応など）を読み込む
. (Join-Path $PSScriptRoot "DeployHelpers.ps1")

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

$sourceDll   = ".\bin\Release\Revit$RevitVersion\Tools28.dll"
$sourcePdb   = ".\bin\Release\Revit$RevitVersion\Tools28.pdb"
$sourceAddin = ".\Packages\$RevitVersion\28Tools\Tools28.addin"
$targetDir      = "$env:ProgramData\Autodesk\Revit\Addins\$RevitVersion"
$targetToolsDir = Get-Tools28TargetDir  $RevitVersion
$pendingDir     = Get-Tools28PendingDir $RevitVersion

# ビルド成果物の確認
if (-not (Test-Path $sourceDll)) {
    Write-Host "エラー: DLLが見つかりません: $sourceDll" -ForegroundColor Red
    Write-Output "DEPLOY_STATUS=FAILED"
    exit 1
}

# .addinファイルの確認
if (-not (Test-Path $sourceAddin)) {
    Write-Host "エラー: .addinファイルが見つかりません: $sourceAddin" -ForegroundColor Red
    Write-Output "DEPLOY_STATUS=FAILED"
    exit 1
}

# ターゲットディレクトリの作成（存在しない場合）
foreach ($d in @($targetDir, $targetToolsDir)) {
    if (-not (Test-Path $d)) {
        Write-Host "ターゲットディレクトリを作成: $d" -ForegroundColor Gray
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}

# Revit の起動状態を確認（メッセージの出し分けに使う）
$revitRunning = Test-RevitRunning
if ($revitRunning) {
    Write-Host "Revit が起動中です。" -ForegroundColor Yellow
    Write-Host "ロックされている DLL は旧ファイルを退避してから差し替えます。" -ForegroundColor Yellow
    Write-Host "（起動中の Revit は古い DLL のまま動き続け、次回起動時に新しい DLL が読み込まれます）" -ForegroundColor Yellow
    Write-Host ""
}

# 前回 Revit 起動中で保留になっていたファイルがあれば先に適用する
$pendingResult = Invoke-Tools28PendingDeploy -RevitVersion $RevitVersion
if ($pendingResult.Applied -gt 0) {
    Write-Host "✓ 保留していた $($pendingResult.Applied) 個のファイルを適用しました" -ForegroundColor Green
}

# デプロイ実行
Write-Host "デプロイ中..." -ForegroundColor Yellow
Write-Host ""

# 配置対象: 全DLL（Tools28.dll + ClosedXML等の依存ライブラリ）＋ PDB
$buildOutputDir = ".\bin\Release\Revit$RevitVersion\"
$deployFiles = @(Get-ChildItem -Path $buildOutputDir -Filter "*.dll" -File)
if (Test-Path $sourcePdb) { $deployFiles += Get-Item $sourcePdb }

$copied   = @()  # そのまま反映された
$replaced = @()  # 差し替え済み（Revit 再起動で反映）
$deferred = @()  # 保留（Revit 終了後に自動適用）
$failed   = @()  # 失敗

foreach ($file in $deployFiles) {
    $result = Copy-Tools28File -SourcePath $file.FullName -TargetDir $targetToolsDir -PendingDir $pendingDir
    switch ($result) {
        "Copied" {
            $copied += $file.Name
            Write-Host "✓ $($file.Name) をコピーしました" -ForegroundColor Green
        }
        "Replaced" {
            $replaced += $file.Name
            Write-Host "✓ $($file.Name) を差し替えました（Revit 再起動で反映）" -ForegroundColor Cyan
        }
        "Pending" {
            $deferred += $file.Name
            Write-Host "⏳ $($file.Name) は保留しました（Revit 終了後に自動適用）" -ForegroundColor Yellow
        }
        default {
            $failed += $file.Name
            Write-Host "✗ $($file.Name) の配置に失敗しました" -ForegroundColor Red
        }
    }
}

# .addinファイルをAddinsルートにコピー（Revit はロックしないので通常コピーでよい）
try {
    Copy-Item $sourceAddin (Join-Path $targetDir "Tools28.addin") -Force -ErrorAction Stop
    Write-Host "✓ Tools28.addin をコピーしました" -ForegroundColor Green
} catch {
    $failed += "Tools28.addin"
    Write-Host "✗ Tools28.addin のコピーに失敗しました" -ForegroundColor Red
}

# 旧バージョンのクリーンアップ（ルートに直置きされていた場合）
$oldRootDll = Join-Path $targetDir "Tools28.dll"
if (Test-Path $oldRootDll) {
    try {
        Remove-Item $oldRootDll -Force -ErrorAction Stop
        Write-Host "✓ 旧 Tools28.dll（ルート直置き）を削除しました" -ForegroundColor Yellow
    } catch {
        Write-Host "⚠ 旧 Tools28.dll（ルート直置き）は使用中のため削除できませんでした" -ForegroundColor Yellow
    }
}

# ========================================
# ステップ3: 反映状況の判定
# ========================================
#
#   APPLIED          … その場で反映済み（Revit を起動すればすぐ使える）
#   RESTART_REQUIRED … 新しい DLL は配置済み。Revit 再起動で反映される
#   PENDING          … 差し替えできず保留中。Revit 終了後に自動で適用される
#   FAILED           … 配置に失敗した

if ($failed.Count -gt 0) {
    $deployStatus = "FAILED"
} elseif ($deferred.Count -gt 0 -or $pendingResult.Remaining -gt 0) {
    $deployStatus = "PENDING"
} elseif ($replaced.Count -gt 0) {
    $deployStatus = "RESTART_REQUIRED"
} else {
    $deployStatus = "APPLIED"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan

switch ($deployStatus) {
    "APPLIED" {
        Write-Host "完了しました！" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "デプロイ先: $targetToolsDir" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Revit は起動していなかったため、そのまま反映済みです。" -ForegroundColor Green
        Write-Host "次のステップ:" -ForegroundColor Yellow
        Write-Host "  1. Revit $RevitVersion を起動してください" -ForegroundColor White
        Write-Host "  2. リボンに「28 Tools」タブが表示されることを確認" -ForegroundColor White
        Write-Host "  3. 機能をテストしてください" -ForegroundColor White
    }
    "RESTART_REQUIRED" {
        Write-Host "完了しました！（Revit 再起動で反映）" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "デプロイ先: $targetToolsDir" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Revit が起動中だったため、新しい DLL に差し替えて配置しました。" -ForegroundColor Cyan
        Write-Host "起動中の Revit には反映されません。Revit $RevitVersion を再起動してください。" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "差し替えたファイル: $($replaced -join ', ')" -ForegroundColor Gray
    }
    "PENDING" {
        Write-Host "ビルドは成功しました（反映待ち）" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Revit が使用中で差し替えられなかったファイルがあります。" -ForegroundColor Yellow
        Write-Host "Revit を終了すると自動で適用され、その後 Revit を起動すると反映されます。" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "保留フォルダ: $pendingDir" -ForegroundColor Gray
    }
    default {
        Write-Host "デプロイに失敗しました" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "失敗したファイル: $($failed -join ', ')" -ForegroundColor Red
    }
}

Write-Host ""
# AutoBuild.ps1 がこの行を読んで通知メッセージを切り替える（機械可読）
# ※ Write-Host は PowerShell 5.1 では「情報ストリーム」に出るため
#   呼び出し元の `$output = & .\QuickBuild.ps1` では拾えない。
#   必ず Write-Output（成功ストリーム）を使うこと。
Write-Output "DEPLOY_STATUS=$deployStatus"
Write-Output "REVIT_RUNNING=$(if ($revitRunning) { 'YES' } else { 'NO' })"

if ($deployStatus -eq "FAILED") { exit 1 }
exit 0
