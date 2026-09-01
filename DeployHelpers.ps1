# Tools28 - デプロイ共通処理 (DeployHelpers.ps1)
#
# QuickBuild.ps1 / AutoBuild.ps1 から dot-source して使う共通関数。
#   . (Join-Path $PSScriptRoot "DeployHelpers.ps1")
#
# ============================================================
# このファイルが解決している問題
# ============================================================
# Revit を起動したままビルドすると、Revit がアドインの DLL
#   C:\ProgramData\Autodesk\Revit\Addins\20XX\28Tools\Tools28.dll
# を掴んだままなので、上書きコピーが必ず失敗する。
# 以前はここでデプロイが中断していたため、あとから Revit を再起動しても
# 古い DLL のままで、ビルドした内容が反映されなかった。
#
# ============================================================
# 対策（3段構え）
# ============================================================
# (1) 通常コピー
#     Revit が起動していない、またはそのファイルがロックされていない場合。
#     その場で反映される。
#
# (2) リネームして差し替え（Revit 起動中の本命）
#     Windows は「読み込み中の DLL」でも "別名へのリネーム" は許可される。
#     そこで旧 DLL を *.old にリネームして退けてから、新しい DLL を
#     本来のファイル名で置く。
#     起動中の Revit はメモリ上の旧 DLL を使い続けるので落ちない。
#     次に Revit を起動したときには新しい DLL が読み込まれる。
#
# (3) 保留フォルダで待機（(2) も失敗した場合の保険）
#     C:\ProgramData\Tools28\PendingDeploy\20XX\ に置いておき、
#     Revit が終了したタイミングで AutoBuild が自動的に適用する。
#
# ※ コンソール出力は英語。PowerShell 5.1 が .ps1 をシステムロケール
#   (Shift-JIS) で読んで日本語が化けるのを避けるため。
#   ユーザー向けの日本語メッセージは AutoBuild.messages.json 側にある。

# 配置先: Revit がアドイン DLL を読み込むフォルダ
function Get-Tools28TargetDir {
    param([Parameter(Mandatory=$true)][string]$RevitVersion)
    return (Join-Path $env:ProgramData "Autodesk\Revit\Addins\$RevitVersion\28Tools")
}

# 保留フォルダ: Revit 起動中で差し替えられなかったファイルの一時置き場
function Get-Tools28PendingDir {
    param([Parameter(Mandatory=$true)][string]$RevitVersion)
    return (Join-Path $env:ProgramData "Tools28\PendingDeploy\$RevitVersion")
}

# Revit が起動中かどうか
function Test-RevitRunning {
    return (@(Get-Process -Name "Revit" -ErrorAction SilentlyContinue).Count -gt 0)
}

# リネーム退避した *.old を掃除する（まだロック中のものは次回に回す）
function Remove-Tools28OldBackups {
    param([Parameter(Mandatory=$true)][string]$TargetDir)

    $removed = 0
    if (-not (Test-Path $TargetDir)) { return $removed }

    Get-ChildItem -Path $TargetDir -Filter "*.old" -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                Remove-Item $_.FullName -Force -ErrorAction Stop
                $removed++
            } catch {
                # まだ Revit が掴んでいる。次回の掃除に任せる
            }
        }
    return $removed
}

# ファイルを1つ配置する
#
# 戻り値:
#   "Copied"   … そのまま上書きできた（すぐ反映）
#   "Replaced" … ロックされていたので旧ファイルをリネーム退避して差し替えた
#                （Revit 再起動で反映）
#   "Pending"  … 差し替えられず保留フォルダに置いた（Revit 終了後に自動適用）
#   "Failed"   … いずれも失敗
function Copy-Tools28File {
    param(
        [Parameter(Mandatory=$true)][string]$SourcePath,
        [Parameter(Mandatory=$true)][string]$TargetDir,
        [string]$PendingDir = ""
    )

    $name = Split-Path $SourcePath -Leaf
    $targetPath = Join-Path $TargetDir $name

    if (-not (Test-Path $TargetDir)) {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }

    # (1) 通常コピー
    try {
        Copy-Item -LiteralPath $SourcePath -Destination $targetPath -Force -ErrorAction Stop
        return "Copied"
    } catch {
        # ロックされている → (2) へ
    }

    # (2) 旧ファイルを *.old にリネームして退避 → 新ファイルを置く
    if (Test-Path $targetPath) {
        $backupName = "$name.$(Get-Date -Format 'yyyyMMdd-HHmmss').old"
        $backupPath = Join-Path $TargetDir $backupName
        $renamed = $false
        try {
            Rename-Item -LiteralPath $targetPath -NewName $backupName -Force -ErrorAction Stop
            $renamed = $true
        } catch {
            # リネームもできなかった → (3) へ
        }

        if ($renamed) {
            try {
                Copy-Item -LiteralPath $SourcePath -Destination $targetPath -Force -ErrorAction Stop
                return "Replaced"
            } catch {
                # 新ファイルを置けなかったので、退避した旧ファイルを元に戻す
                # （これをしないとアドインが読み込めなくなる）
                try { Rename-Item -LiteralPath $backupPath -NewName $name -Force -ErrorAction Stop } catch { }
            }
        }
    }

    # (3) 保留フォルダに置いて、Revit 終了後に適用する
    if ($PendingDir) {
        try {
            if (-not (Test-Path $PendingDir)) {
                New-Item -ItemType Directory -Path $PendingDir -Force | Out-Null
            }
            Copy-Item -LiteralPath $SourcePath -Destination (Join-Path $PendingDir $name) -Force -ErrorAction Stop
            return "Pending"
        } catch {
            return "Failed"
        }
    }

    return "Failed"
}

# 保留フォルダにたまっているファイルを配置先へ適用する
#
# 戻り値: PSCustomObject
#   Applied   … 適用できたファイル数
#   Remaining … まだ適用できず残っているファイル数
function Invoke-Tools28PendingDeploy {
    param([Parameter(Mandatory=$true)][string]$RevitVersion)

    $targetDir  = Get-Tools28TargetDir  $RevitVersion
    $pendingDir = Get-Tools28PendingDir $RevitVersion

    $applied   = 0
    $remaining = 0

    if (Test-Path $pendingDir) {
        Get-ChildItem -Path $pendingDir -File -ErrorAction SilentlyContinue | ForEach-Object {
            # $PendingDir は渡さない（保留フォルダから保留フォルダへのコピーを防ぐ）
            $result = Copy-Tools28File -SourcePath $_.FullName -TargetDir $targetDir
            if ($result -eq "Copied" -or $result -eq "Replaced") {
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                $applied++
            } else {
                $remaining++
            }
        }
    }

    # 退避済みの *.old を掃除（Revit 終了後ならここで消える）
    Remove-Tools28OldBackups -TargetDir $targetDir | Out-Null

    return [PSCustomObject]@{
        Applied   = $applied
        Remaining = $remaining
    }
}
