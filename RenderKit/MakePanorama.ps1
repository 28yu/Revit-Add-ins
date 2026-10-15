<#
.SYNOPSIS
    Revit から書き出した FBX を Blender で 360°パノラマにし、ブラウザで開ける HTML を作る。

.DESCRIPTION
    普段は MakePanorama.bat に FBX をドラッグ＆ドロップして使う。
    細かく指定したいときは、このスクリプトを直接実行する。

.EXAMPLE
    .\MakePanorama.ps1 -Model C:\work\house.fbx
.EXAMPLE
    .\MakePanorama.ps1 -Model C:\work\house.fbx -Quality high -SunAzimuth 180 -SunElevation 60
.EXAMPLE
    .\MakePanorama.ps1 -Model C:\work\house.fbx -Eye "2.5,3.0,1.5;8.0,3.0,1.5"
#>
param(
    [Parameter(Mandatory = $true)][string]$Model,
    [ValidateSet("draft", "normal", "high")][string]$Quality = "normal",
    [string[]]$Eye = @(),             # 撮影位置 "X,Y,Z"（メートル）。";" 区切りで複数書くとその数だけ作る
    [double]$EyeHeight = 1.5,         # 撮影位置を自動で決めるときの目の高さ
    [double]$SunAzimuth = 135,        # 太陽の方位（0=北, 90=東, 180=南）
    [double]$SunElevation = 40,       # 太陽の高さ（0=地平線, 90=真上）
    [double]$Exposure = 0,            # 明るさの補正（+1 で2倍、-1 で半分）
    [switch]$NoAutoExposure,          # 自動露出・自動ホワイトバランスを使わない
    [switch]$Cpu,                     # GPU を使わず CPU で計算する
    [switch]$SaveBlend,               # Blender のファイル（.blend）も保存する
    [string]$Blender = ""             # blender.exe の場所（省略時は自動で探す）
)

$ErrorActionPreference = "Stop"
# Blender が出す日本語のメッセージを文字化けさせない
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

function Find-Blender {
    param([string]$Preferred)
    if ($Preferred) {
        if (Test-Path $Preferred) { return (Resolve-Path $Preferred).Path }
        throw "指定された blender.exe が見つかりません: $Preferred"
    }
    if ($env:BLENDER_EXE -and (Test-Path $env:BLENDER_EXE)) { return $env:BLENDER_EXE }

    # 通常のインストール先（複数あれば一番新しいバージョン）
    $roots = @(
        "$env:ProgramFiles\Blender Foundation",
        "${env:ProgramFiles(x86)}\Steam\steamapps\common"
    )
    $found = foreach ($root in $roots) {
        if (Test-Path $root) {
            Get-ChildItem -Path $root -Directory -Filter "Blender*" -ErrorAction SilentlyContinue |
                ForEach-Object { Join-Path $_.FullName "blender.exe" } |
                Where-Object { Test-Path $_ }
        }
    }
    $latest = $found | Sort-Object { (Get-Item $_).VersionInfo.FileVersionRaw } -Descending | Select-Object -First 1
    if ($latest) { return $latest }

    $cmd = Get-Command blender.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

$modelPath = (Resolve-Path $Model).Path
$blenderExe = Find-Blender -Preferred $Blender
if (-not $blenderExe) {
    Write-Host ""
    Write-Host "Blender が見つかりませんでした。" -ForegroundColor Red
    Write-Host "  1. https://www.blender.org/download/ から Blender（4.2 以降）をインストールしてください"
    Write-Host "  2. 別の場所に入れた場合は、環境変数 BLENDER_EXE に blender.exe の場所を設定してください"
    exit 1
}

$script = Join-Path $PSScriptRoot "blender\render_panorama.py"
$outHtml = Join-Path (Split-Path $modelPath -Parent) (([IO.Path]::GetFileNameWithoutExtension($modelPath)) + "_360.html")

# 小数点が「,」になる地域設定でも正しく渡せるよう、数値は「.」区切りの文字にする
function Format-Number([double]$value) { $value.ToString([Globalization.CultureInfo]::InvariantCulture) }

$pyArgs = @(
    "--input", $modelPath,
    "--output", $outHtml,
    "--quality", $Quality,
    "--eye-height", (Format-Number $EyeHeight),
    "--sun-azimuth", (Format-Number $SunAzimuth),
    "--sun-elevation", (Format-Number $SunElevation),
    "--exposure", (Format-Number $Exposure)
)
# "-2,3,1.5" のようにマイナスで始まるとオプション名と間違われるため「--eye=値」の形で渡す
foreach ($e in $Eye) { $pyArgs += "--eye=$e" }
if ($NoAutoExposure) { $pyArgs += "--no-auto-exposure" }
if ($Cpu) { $pyArgs += @("--device", "cpu") }
if ($SaveBlend) { $pyArgs += "--save-blend" }

Write-Host "Blender: $blenderExe"
Write-Host "モデル : $modelPath"
Write-Host "品質   : $Quality"
Write-Host ""

# -b: 画面を開かずに実行 / --python-exit-code 1: スクリプトが失敗したらエラーで終わる
& $blenderExe -b --python-exit-code 1 -P $script -- @pyArgs
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "失敗しました（終了コード $LASTEXITCODE）。上のメッセージを確認してください。" -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "完成しました: $outHtml" -ForegroundColor Green
try {
    Invoke-Item $outHtml  # 既定のブラウザで開く
} catch {
    Write-Host "（ブラウザを自動で開けませんでした。上のファイルをダブルクリックしてください）"
}
