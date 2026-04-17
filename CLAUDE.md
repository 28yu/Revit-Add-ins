# Tools28 - Revit Add-in 開発ガイド (Claude Code用)

## コミットメッセージのルール
- **コミットメッセージは必ず日本語で書くこと**
- 1行目: 変更内容の要約（日本語、50文字以内目安）
- 例: `通知メッセージを日本語化`, `梁天端色分け機能を追加`, `ビルドスクリプトの権限エラーを修正`

## ⚠️ 多言語化ルール（必須）
- **新機能の追加・既存機能の改善時は、必ず3言語（JP/US/CN）を同時に対応すること**
- UI上に表示される全ての文字列（ボタン名、ツールチップ、ダイアログテキスト、エラーメッセージ、TaskDialog等）は `Loc.S("Key")` を使用し、ハードコードされた日本語文字列を書かない
- 翻訳エントリは `Localization/StringsJP.cs`, `StringsEN.cs`, `StringsCN.cs` の3ファイルに同じキーで同時追加
- リボンボタン追加時は `_buttonTextKeys` / `_buttonTipKeys` へのマッピング追加も忘れないこと
- 詳細は「新機能追加手順 > 2. 多言語リソースの追加」を参照

## プロジェクト概要
- **名前**: Tools28
- **種類**: Autodesk Revit アドイン (C# / .NET Framework 4.8)
- **対応バージョン**: Revit 2021, 2022, 2023, 2024, 2025, 2026
- **名前空間**: `Tools28`
- **アセンブリ名**: `Tools28`
- **プロジェクトファイル**: `Tools28.csproj` (SDK-style, マルチバージョン対応)

## リポジトリ構成

```
Revit-Add-ins/
├── Application.cs              # メインアプリ (IExternalApplication) - リボンUI構築
├── Tools28.csproj              # SDK-style マルチバージョン対応プロジェクトファイル
├── dev-config.json             # 開発設定（開発用Revitバージョン指定）
├── .claude/
│   └── settings.json           # Claude Code設定 (push前自動rebase hook等)
├── Commands/                   # 機能コマンド群
│   ├── GridBubble/             # 通り芯・レベルの符号表示切替
│   ├── SheetCreation/          # シート一括作成 (WPFダイアログ付き)
│   ├── ViewCopy/               # 3Dビュー視点コピー
│   ├── SectionBoxCopy/         # セクションボックスコピー
│   ├── ViewportPosition/       # ビューポート位置コピー (自動マッチング)
│   ├── CropBoxCopy/            # トリミング領域コピー
│   ├── BeamTopLevel/           # 梁天端レベル色分け (平面ビュー)
│   └── LanguageSwitch/         # 言語切替・バージョン情報・マニュアル
├── Localization/               # 多言語リソース
│   ├── Loc.cs                  # 多言語マネージャー (静的クラス)
│   ├── StringsJP.cs            # 日本語文字列辞書
│   ├── StringsEN.cs            # 英語文字列辞書
│   └── StringsCN.cs            # 中国語（簡体字）文字列辞書
├── Resources/Icons/            # 32x32アイコン (12個)
├── Properties/                 # AssemblyInfo.cs
├── Packages/                   # 配布パッケージテンプレート (バージョン別)
│   ├── 2021/ ~ 2026/           # 各バージョン用
│   │   ├── 28Tools/            #   Tools28.addin (DLLはビルド時にコピー)
│   │   ├── install.bat         #   自動インストール
│   │   ├── uninstall.bat       #   アンインストール
│   │   └── README.txt          #   インストール手順
├── .github/workflows/          # GitHub Actions
│   ├── build-and-release.yml   #   タグ push or 手動実行で配布ZIP生成
│   ├── auto-merge-claude.yml   #   claude/* ブランチの自動PR・マージ
│   └── deploy-pages.yml        #   Docs/ を GitHub Pages にデプロイ
├── Dist/                       # 配布ZIP出力先 (git管理外)
├── QuickBuild.ps1              # 🚀 高速ビルド＆デプロイ（開発用）
├── BuildAll.ps1                # 全バージョン一括ビルド（リリース用）
├── GenerateAddins.ps1          # .addinマニフェスト生成
├── CreatePackages.ps1          # 配布ZIP作成
├── Docs/                       # GitHub Pages (ワークフロー図等)
│   ├── workflow-diagram.html   #   開発ワークフロー図
│   └── icons/workflow/         #   ワークフロー図用アイコン
├── Deploy-For-Testing.ps1      # テスト用デプロイ（手動）
├── DEVELOPMENT.md              # 開発者ガイド（詳細手順）
└── CLAUDE.md                   # このファイル
```

## ビルド

### ターゲットフレームワーク (SDK-style csproj)
- **Revit 2021-2024**: `net48` (.NET Framework 4.8)
- **Revit 2025-2026**: `net8.0-windows` (.NET 8)

バージョンに応じて `Tools28.csproj` 内で自動切替え。

### 条件付きコンパイルシンボル
`REVIT2021`, `REVIT2022`, `REVIT2023`, `REVIT2024`, `REVIT2025`, `REVIT2026`

### ビルドコマンド
```powershell
# 全バージョン
.\BuildAll.ps1

# 特定バージョン
msbuild Tools28.csproj /p:Configuration=Release /p:RevitVersion=2024
```

### 出力先
`bin\Release\Revit{VERSION}\Tools28.dll`

## 配布パッケージ

### 配布ZIP構成
配布ZIPのファイル名は `28Tools_Revit{VERSION}_vX.X.zip`。
```
28Tools_Revit{VERSION}_vX.X.zip
├── 28Tools/
│   ├── Tools28.dll              # メインDLL
│   ├── Tools28.addin            # マニフェストファイル
│   ├── ClosedXML.dll            # Excel読み書きライブラリ
│   └── (その他依存DLL)          # ClosedXMLの依存ライブラリ群
├── install.bat                  # 自動インストール
├── uninstall.bat                # アンインストール
└── README.txt                   # インストール手順
```

### テンプレート構成 (Packages/{VERSION}/)
各バージョン (2021-2026) に以下を格納 (DLLはビルド時に自動コピー):
```
Packages/{VERSION}/
├── 28Tools/
│   └── Tools28.addin
├── install.bat
├── uninstall.bat
└── README.txt
```

### install.bat の内容
- `chcp 65001` でUTF-8対応
- バージョン番号をヘッダーに表示（例: `28 Tools v2.0 for Revit 2024`）
- `28Tools\` フォルダ内の全DLLを `Addins\{VERSION}\28Tools\` にコピー
- `Tools28.addin` を `Addins\{VERSION}\` ルートにコピー
- 旧バージョン（ルート直置きDLL）の自動クリーンアップ
- マニュアルURLの表示

### uninstall.bat の内容
- `C:\ProgramData\Autodesk\Revit\Addins\{VERSION}\28Tools\` フォルダを削除
- `Tools28.addin` を削除
- 旧バージョン（ルート直置き `Tools28.dll`）のクリーンアップ

### README.txt の内容
- バージョン番号付きタイトル（例: `28 Tools v2.0 for Revit 2024`）
- クイックスタート手順 (install.bat を管理者実行 → Revit 再起動)
- 新機能・変更点セクション
- 機能一覧（パネル別）
- マニュアルURL
- アンインストール手順
- 対応バージョン

### インストール先
```
C:\ProgramData\Autodesk\Revit\Addins\{VERSION}\
├── Tools28.addin                    # マニフェスト（ルートに配置）
└── 28Tools\                         # DLLサブフォルダ
    ├── Tools28.dll
    ├── ClosedXML.dll
    └── (その他依存DLL)
```

### 配布ZIP作成手順
```powershell
# 1. 全バージョンをビルド
.\BuildAll.ps1

# 2. 配布ZIPを作成 (バージョン番号を指定)
.\CreatePackages.ps1 -Version "2.0"

# 出力先: .\Dist\28Tools_Revit20XX_v2.0.zip
```

### ⚠️ リリース時の配布パッケージ更新チェックリスト

**新機能の追加やバージョン番号の変更時は、以下のファイルを必ず更新すること：**

1. **`Localization/StringsJP.cs`, `StringsEN.cs`, `StringsCN.cs`** (3ファイル同時)
   - 新機能のUI文字列（ボタン名、ツールチップ、ダイアログ文字列、エラーメッセージ等）
   - 3ファイル全てに同じキーを追加（キー不一致は禁止）

2. **`Packages/{VERSION}/README.txt`** (全6バージョン)
   - タイトルのバージョン番号（例: `28 Tools v2.0 for Revit 2024`）
   - 「新機能・変更点」セクションに追加機能を記載
   - 「機能一覧」セクションに新しいコマンドを追加

3. **`Packages/{VERSION}/install.bat`** (全6バージョン)
   - ヘッダーのバージョン番号（例: `28 Tools v2.0 for Revit 2024`）

4. **`.github/workflows/build-and-release.yml`**
   - Release body の「機能一覧」セクションに新機能を追加

5. **`Application.cs`**
   - リボンに新しいボタンを登録
   - `_buttonTextKeys` / `_buttonTipKeys` にマッピングを追加（言語切替時の動的更新用）

#### 更新コマンド例（全6バージョン一括更新）
```bash
# README.txt のバージョン番号を一括置換
for ver in 2021 2022 2023 2024 2025 2026; do
  sed -i 's/v2.0/v2.1/g' Packages/$ver/README.txt
  sed -i 's/v2.0/v2.1/g' Packages/$ver/install.bat
done
```

## 開発ワークフロー

### 日常的な開発サイクル（Revit 2022ベース）

```powershell
# 1. 機能の実装・修正
#    Commands/ 配下にコマンドクラスを作成
#    Application.cs にリボンボタンを登録

# 2. クイックビルド＆デプロイ（Revit 2022のみ）
.\QuickBuild.ps1

# 3. Revit 2022を起動してテスト

# 4. 問題があれば修正して再度 QuickBuild.ps1
```

### リリース準備（完成後）

```powershell
# 1. 全バージョン（2021-2026）をビルド
.\BuildAll.ps1

# 2. 配布ZIPを作成
.\CreatePackages.ps1 -Version "1.1"

# 3. 動作確認（必要に応じて複数バージョンで検証）

# 4. コミット＆プッシュ
git add .
git commit -m "Add new feature"
git push -u origin claude/setup-addon-workflow-yO1Uz

# 5. GitHub Releasesで公開（自動）
git tag v1.1
git push --tags
```

## 新機能追加手順

### 1. コマンドクラスの作成

`Commands/` に新しいフォルダを作成し、`IExternalCommand` を実装：

```csharp
// Commands/FeatureName/FeatureNameCommand.cs
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.FeatureName
{
    [Transaction(TransactionMode.Manual)]
    public class FeatureNameCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // 実装
            return Result.Succeeded;
        }
    }
}
```

### 2. 多言語リソースの追加（必須）

**新機能追加・既存機能改善時は、必ず3言語同時に対応すること。**

#### 2-1. 翻訳辞書にエントリを追加

3ファイル全てに同じキーを追加する（キーの不一致は禁止）：

```csharp
// Localization/StringsJP.cs
{ "FeatureName.Title", "機能名タイトル" },
{ "FeatureName.Description", "機能の説明文" },

// Localization/StringsEN.cs
{ "FeatureName.Title", "Feature Name Title" },
{ "FeatureName.Description", "Description of the feature" },

// Localization/StringsCN.cs
{ "FeatureName.Title", "功能名称标题" },
{ "FeatureName.Description", "功能说明" },
```

キー命名規則:
- リボンボタン: `Ribbon.FeatureName.ButtonName` / `Ribbon.FeatureName.ButtonName.Tip`
- パネル名: `Ribbon.Panel.PanelName`
- ダイアログ: `FeatureName.Title`, `FeatureName.LabelX`
- 共通: `Common.OK`, `Common.Cancel` 等は既存を再利用

#### 2-2. コード内の文字列を `Loc.S()` で置換

```csharp
// ❌ ハードコードされた日本語
btn.ToolTip = "機能の説明";
TaskDialog.Show("エラー", "選択してください");

// ✅ 多言語対応
btn.ToolTip = Loc.S("Ribbon.FeatureName.Tip");
TaskDialog.Show(Loc.S("Common.Error"), Loc.S("FeatureName.SelectPrompt"));
```

#### 2-3. WPFダイアログがある場合

`ApplyLocalization()` メソッドを実装し、コンストラクタで呼び出す：

```csharp
private void ApplyLocalization()
{
    this.Title = Loc.S("FeatureName.Title");
    labelX.Text = Loc.S("FeatureName.LabelX");
    btnOK.Content = Loc.S("Common.OK");
    btnCancel.Content = Loc.S("Common.Cancel");
}
```

### 3. リボンへの登録

`Application.cs` の `OnStartup()` メソッド内でボタンを追加し、多言語マッピングも登録：

```csharp
// ボタン作成（テキストとツールチップは Loc.S() で取得）
PushButton btn = panel.AddItem(new PushButtonData(
    "FeatureName",
    Loc.S("Ribbon.FeatureName.Button"),
    assemblyPath,
    "Tools28.Commands.FeatureName.FeatureNameCommand"
)) as PushButton;
btn.ToolTip = Loc.S("Ribbon.FeatureName.Button.Tip");

// 言語切替時の動的更新用マッピングに追加
_buttons["FeatureName"] = btn;
_buttonTextKeys["FeatureName"] = "Ribbon.FeatureName.Button";
_buttonTipKeys["FeatureName"] = "Ribbon.FeatureName.Button.Tip";
```

### 5. アイコンの追加（オプション）

`Resources/Icons/` に32x32 PNGを追加し、`.csproj` に `<Resource>` を追加

```xml
<ItemGroup>
  <Resource Include="Resources\Icons\FeatureName.png" />
</ItemGroup>
```

### 6. ビルド＆テスト

```powershell
.\QuickBuild.ps1  # Revit 2022でビルド→デプロイ
```

※ SDK-style csproj のため `.cs` ファイルは自動認識される（`<Compile Include>` は不要）

## 外部参照

- **マニュアル**: https://28tools.com/addins.html
- **配布サイト**: https://28yu.github.io/28tools-download/
- **リポジトリ**: https://github.com/28yu/Revit-Add-ins

## CI/CD (GitHub Actions)

### 自動リリース (タグ push)
```bash
git tag v1.0
git push --tags
```
GitHub Actions が自動的に:
1. 全6バージョン (2021-2026) をビルド
2. 配布ZIP (`28Tools_Revit20XX_v1.0.zip`) を作成
3. GitHub Releases にアップロード

### 手動実行
GitHub → Actions → "Build and Release" → Run workflow → バージョン番号を入力

### ワークフロー実行タイトル (run-name)
- 各ワークフローに `run-name:` を設定し、Actions 一覧の太字タイトルを日本語化
- `name:` はサイドバーのワークフロー名（英語のまま）、`run-name:` は各実行のタイトル
- auto-merge: コミットメッセージをそのまま表示（`${{ github.event.head_commit.message }}`）

### Revit API 参照
NuGet パッケージ `Nice3point.Revit.Api` を使用 (ローカル Revit 不要)

### 自動マージ＆デプロイ (`claude/**` ブランチ)

`claude/` で始まるブランチを push すると、以下が自動実行される：

1. **auto-merge-claude.yml**: main へ直接 squash merge → ブランチ削除（PR不要）
2. **deploy-pages.yml**: main への push / merge をトリガーに `Docs/` を GitHub Pages へデプロイ
3. **ローカル自動ビルド**: `AutoBuild.ps1` が main の更新を検知 → 自動 pull → ビルド → デプロイ → Windows 通知

#### ⚠️ push 前の rebase（PreToolUse hookで自動化済み）

自動マージはコンフリクトがあると失敗する。**push 前に必ず rebase が必要**だが、
`.claude/settings.json` の PreToolUse hook により **`git push` 実行前に自動的に
`git fetch origin main && git rebase origin/main` が実行される**ため、手動での rebase は不要。

**hook の仕組み**:
- Bash ツールで `git push` を含むコマンドが実行される直前にフックが発動
- 未コミット変更がある場合は `git stash` で退避してからリベース
- `git fetch origin main` で最新の main を取得
- `git rebase origin/main` で現在のブランチを main の先頭にリベース
- 退避した変更がある場合は `git stash pop` で復元
- リベース完了後に push が実行されるため、コンフリクトによる自動マージ失敗を防止

**⚠️ squash merge によるオートマージ連続失敗の既知問題と対策**:

自動マージは **squash merge** で行われるため、main 上のコミットハッシュはブランチのものと異なる。
連続して push → 自動マージ → push を繰り返すと、ブランチに「既にmainに含まれる変更」の
旧コミットが残り、`git merge --squash` がコンフリクトで失敗することがある。

**根本原因**:
1. push → squash merge 成功 → main に「Auto: xxx」コミットが作成される
2. 次の変更を push → ブランチには旧コミット + 新コミットが乗っている
3. `git merge --squash` が旧コミットの変更と main の squash コミットでコンフリクト

**対策（hookで自動化済み）**:
- PreToolUse hook が push 前に `git rebase origin/main` を自動実行
- rebase により旧コミットは `skipped previously applied commit` としてスキップされる
- 新コミットのみがブランチに残るため、auto-merge が成功する
- hook は dirty working tree も `git stash` で対応済み

**⚠️ commit と push は必ず別の Bash 呼び出しで実行すること**:
- `git commit && git push` を1つの Bash 呼び出しにすると、hook が push 前に rebase を
  実行する時点でまだ commit が完了していない場合がある
- 必ず `git commit` と `git push` を **別々の Bash ツール呼び出し** にすること

**手動で rebase が必要な場合**:
- rebase 中にコンフリクトが発生した場合は hook が失敗するため、手動で解決が必要
- `git rebase --continue` でコンフリクト解決後、再度 push する

特に同じファイル（例: `Docs/workflow-diagram.html`）を連続して変更する場合は要注意。

#### AutoBuild（バックグラウンド自動ビルド＆デプロイ）

**起動方法**: `StartAutoBuild.vbs` をダブルクリック → UAC「はい」→ バックグラウンドで常駐

**動作フロー**:
1. 30秒間隔で `origin/main` を `git fetch`
2. 変更検知 → `git reset --hard origin/main` → `QuickBuild.ps1` 実行
3. ビルド成功 → `C:\ProgramData` にデプロイ → 日本語通知ダイアログ表示

**ファイル構成**:
- `StartAutoBuild.vbs` — VBS ラッパー（`Shell.Application.ShellExecute` + `runas` で管理者昇格）
- `AutoBuild.ps1` — メインスクリプト（監視ループ、ビルド、通知）
- `AutoBuild.log` — メインログ（リポジトリ直下）
- `AutoBuild_detail.log` — ビルド詳細出力（リポジトリ直下）

**⚠️ AutoBuild.ps1 自体を変更した場合の再起動手順**:
1. タスクマネージャー → 詳細タブ → AutoBuild の `powershell.exe` を終了（管理者権限で動作中のため、通常の `Stop-Process` では権限エラー）
2. `git pull origin main` で最新を取得
3. `StartAutoBuild.vbs` をダブルクリック → UAC「はい」

#### 開発で得た知見（AutoBuild）

##### 管理者権限
- `C:\ProgramData\Autodesk\Revit\Addins\` への書き込みには管理者権限が必要
- VBS の `ShellExecute "runas"` で UAC 昇格 → 管理者権限の PowerShell を起動
- 昇格後のプロセスは `C:\Windows\System32` がカレントディレクトリになるため、`Set-Location $PSScriptRoot` が必須

##### PowerShell 5.1 での日本語文字化け
- Windows PowerShell 5.1 は `.ps1` ファイルをシステムロケール（Shift-JIS）で読む
- UTF-8 BOM を付けても `git reset --hard` 後に失われる場合がある
- **解決策**: 日本語文字列は Unicode エスケープで記述する
  ```powershell
  # "ビルド成功" を Unicode エスケープで
  $msg = -join([char[]]@(0x30D3,0x30EB,0x30C9,0x6210,0x529F))
  ```
- 通知ダイアログへの日本語テキスト受け渡しは JSON ファイル経由 + `[System.IO.File]::WriteAllText/ReadAllText` で UTF-8 を明示指定

##### 通知ダイアログ (MessageBox)
- `Start-Process powershell -Command` で日本語を渡すと文字化け
- `-EncodedCommand`（Base64）でも日本語を含むスクリプトは失敗
- **解決策**: 日本語テキストを JSON ファイルに書き出し、`-EncodedCommand` のスクリプトは JSON を読むだけ（ASCII のみ）にする

##### 外部コマンド (git) の日本語出力
- `git log` 等の出力は UTF-8 だが、PowerShell 5.1 はシステムロケール（Shift-JIS）で読み取る
- **解決策**: スクリプト冒頭で `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8` を設定

##### ビルド成功判定
- `& .\QuickBuild.ps1` の `$LASTEXITCODE` は信頼できない（PowerShell スクリプト呼び出しでは正しく伝搬しない）
- **解決策**: ビルド前後の DLL タイムスタンプを比較して成功判定

#### ⚠️ ローカルでの pull 失敗に注意

Claude Code で push → 自動マージ後、ユーザーがローカルで `git pull origin main` する際：
- **ローカルに未コミット変更があると `git pull` がエラーで中止される**
- エラーメッセージを見落とすと、古いコードのままビルドしてしまう
- 特にバイナリファイル（PNG アイコン等）は Claude Code とローカルの両方で変更されやすい
- **対策**: pull 後に `git log --oneline -3` で期待するコミットが反映されたか確認する

## GitHub Pages (Docs/)

### 概要
`Docs/` フォルダが GitHub Pages としてデプロイされる。

- **URL**: `https://28yu.github.io/Revit-Add-ins/`
- **ワークフロー図**: `Docs/workflow-diagram.html`
- **アイコン**: `Docs/icons/workflow/` (you.png, claude.png, github.png 等)

### ワークフロー図の配色 (現在)

| 色 | コード | 用途 |
|---|---|---|
| Blue-green | `#5F968E` | フェーズヘッダー、矢印 |
| Mint | `#BFDCCF` | あなた（You）ノード背景 |
| Oatmeal | `#D5C9B1` | AI・自動（Auto）ノード背景、結果ノード、ブラウザフレーム |

派生色（背景色・ボーダー等）はメインカラーから明度を調整して生成。

## 注意事項

- Revit API は NuGet パッケージ経由で取得 (`Nice3point.Revit.Api.RevitAPI` / `RevitAPIUI`)
- トランザクションは `TransactionMode.Manual` を使用
- デバッグログは `C:\temp\Tools28_debug.txt` に出力
- WPFダイアログを使用するコマンドは XAML + コードビハインドで構成

## Revit API の既知の制限事項

- **ラインワークツール (Linework Tool) のオーバーライドは API で取得・設定不可** — `Edge.GraphicsStyleId` はビュー固有のラインワーク変更を反映しない。`IExportContext2D` でも同様。Autodesk が公式に API ギャップとして認めている。詳細は `DEVELOPMENT.md` の「Revit API の既知の制限事項」セクションを参照。

## BeamUnderLevel（梁下端色分け）設計メモ

### 概要
天井伏図上の梁を下端レベル値で色分けするコマンド。

### 計算式
```
梁下端レベル = 階高 - 梁天端レベル - 梁高さ
```
コード上: `bottomLevel = floorHeight + topLevelOffset - beamHeight`
- `floorHeight` = 上位レベル標高 − 参照レベル標高（例: 3000mm）
- `topLevelOffset` = 梁天端パラメータ値（上位レベルからの下がりは負値。例: -300mm）
- `beamHeight` = 梁高さパラメータ値（例: 600mm）
- 結果は参照レベル基準（例: +2100mm → 参照レベルから2100mm上）

### レベル構成
- **参照レベル**: 天井伏図の GenLevel（自動取得、変更不可）
- **上位レベル**: ユーザーが選択（参照レベルより上のレベルのみ表示）
- 梁は上位レベルに配置され、オフセット（通常負値）で下がっている前提

### ダイアログ構成（4ステップ）
1. レベル設定（参照レベル表示 + 上位レベル選択 + 階高表示）
2. 梁高さパラメータ選択（ファミリ毎）
3. 梁天端レベルパラメータ選択（ファミリ毎）
4. 処理確認・実行

### パラメータ選択の設計
- ファミリ毎に異なるパラメータを選択可能
- 主要候補はラジオボタン（自動検出、検出数表示）
- 「その他」はComboBox（レベル・オフセット関連キーワードでフィルタしたパラメータ一覧）
- テキスト入力のカスタム欄は不要（ComboBoxで十分）

### フィルタ・色分けの設計
- **グラフィック上書き**: 投影サーフェス前景の塗り潰しのみ（断面パターン・投影線・断面線は変更しない）
- **配色**: 明るいパステル〜中間色トーンのみ使用。黒っぽい色・暗い茶色は使わない
- **フィルタ名**: `梁下_{レベル名}{±値}` 形式
- サーフェスを持たない梁ファミリはファミリ側でサーフェス表示を有効にする（アドイン側では対応不要）

### コード構成
```
Commands/BeamUnderLevel/
├── BeamUnderLevelCommand.cs    # メインコマンド (IExternalCommand)
├── BeamUnderLevelDialog.xaml    # WPF 4ステップダイアログ
├── BeamUnderLevelDialog.xaml.cs # ダイアログのコードビハインド
├── BeamUnderLevelDialogData.cs  # ダイアログのデータモデル
├── BeamCalculator.cs            # 梁下端レベル計算ロジック
├── FilterManager.cs             # ParameterFilterElement の作成・色分け適用
├── ParameterManager.cs          # 共有パラメータ (梁下端レベル) の作成・値設定
├── LegendManager.cs             # 凡例ビュー作成（色分け対応表）
└── BeamLabelManager.cs          # 梁上に下端レベル値の TextNote を配置
```

### 開発で得た知見

#### アイコン作成
- アイコンは `Resources/Icons/{name}_32.png` の命名規則
- `Tools28.csproj` の `<Resource>` に登録が必要
- `Application.cs` の `LoadImage()` でリソースまたはファイルから読み込み（ハイブリッド方式）

#### 梁ラベル (TextNote) の配置
- ビュー上の梁の位置取得には `beam.get_BoundingBox(view)` を使用（モデル座標の BoundingBox ではなくビュー固有のものを使うこと）
- 梁の幅はタイプパラメータから取得（インスタンスパラメータではない）
- ラベルのオフセット量はビュースケール (`view.Scale`) を考慮して調整する
- テキスト配置は `Center` + `Bottom` で梁との重なりを防止
- 梁の方向に合わせてラベルを回転させる

#### 自動マージ (claude/* ブランチ)
- push 前の rebase は PreToolUse hook で自動実行される（`.claude/settings.json`）
- 自動マージ成功後はリモートブランチが自動削除される
- 削除後に再 push する場合は `--force-with-lease` ではなく通常の push を使用
- マージ失敗時は rebase して再 push すれば自動リトライされる

### 現在のステータス
- **機能実装**: 完了（ダイアログ、計算、フィルタ、ラベル、凡例）
- **動作確認**: Revit 環境で動作確認済み
- **アイコン**: I型梁 + 上向き矢印 + ∇FL線 + 3色ブロック（`beam_under_level_32.png`）

## BeamTopLevel（梁天端色分け）設計メモ

### 概要
平面ビュー上の梁を天端レベル値で色分けするコマンド。BeamUnderLevel をベースに簡略化。

### 計算式
計算式は不要。梁の天端レベルパラメータの値（参照レベル基準）をそのまま使用する。

### 対象ビュー
- **平面ビュー（FloorPlan）および構造伏図（EngineeringPlan）**
- 参照レベルはビューの GenLevel から自動取得

### BeamUnderLevel との違い
| 項目 | BeamUnderLevel | BeamTopLevel |
|------|---------------|-------------|
| 対象ビュー | 天井伏図 (CeilingPlan) | 平面ビュー (FloorPlan) / 構造伏図 (EngineeringPlan) |
| 計算式 | 階高 + オフセット - 梁高さ | パラメータ値をそのまま使用 |
| ダイアログ | 4ステップ | 3ステップ |
| レベル設定 | 上位レベル選択あり | 不要（参照レベル自動取得） |
| 梁高さパラメータ | 選択必要 | 不要 |
| プレフィックス | 梁下_ | 梁天端_ |
| ラベルサフィックス | 下端 | 天端 |

### ダイアログ構成（3ステップ）
1. 基本設定（ビュー情報表示 + 文字タイプ選択）
2. 梁天端レベルパラメータ選択（ファミリ毎）
3. 処理確認・実行

### フィルタ・色分けの設計
- BeamUnderLevel と同じ設計（投影サーフェス前景の塗り潰しのみ）
- フィルタ名: `梁天端_{レベル名}{±値}` 形式
- 共有パラメータ: `梁天端_基準レベル`, `梁天端_レベル差`, `梁天端_表示`, `梁天端_エラー`

### コード構成
```
Commands/BeamTopLevel/
├── BeamTopLevelCommand.cs      # メインコマンド (IExternalCommand)
├── BeamTopLevelDialog.xaml      # WPF 3ステップダイアログ
├── BeamTopLevelDialog.xaml.cs   # ダイアログのコードビハインド
├── BeamTopLevelDialogData.cs    # ダイアログのデータモデル
├── BeamCalculator.cs            # 天端レベル取得ロジック（パラメータ値直接取得）
├── FilterManager.cs             # ParameterFilterElement の作成・色分け適用
├── ParameterManager.cs          # 共有パラメータ (梁天端レベル) の作成・値設定
├── LegendManager.cs             # 凡例ビュー作成（色分け対応表）
└── BeamLabelManager.cs          # 梁上に天端レベル値の TextNote を配置
```

### ダイアログの設計知見
- `SizeToContent="Height"` + `MaxHeight="800"` でコンテンツに応じた自動サイズ調整（固定Heightだと隙間が生じる）
- Step1のGrid行定義で `Height="*"` を使うと不要な空間ができるため `Height="Auto"` のみにする

### アイコンのデザイン規則
- 梁下端: I型梁(上) + 上向き矢印 + ∇FL線(下) + ピンク/黄/青3色ブロック(右)
- 梁天端: ∇FL線(上) + 下向き矢印 + I型梁(下) + ピンク/黄/青3色ブロック(右)
- 色: ピンク `(255,128,148)`, 黄 `(218,185,47)`, 青 `(30,144,255)`
- Python Pillow (`ImageDraw`) で32x32 PNGを生成

### 現在のステータス
- **機能実装**: 完了（ダイアログ、パラメータ取得、フィルタ、ラベル、凡例）
- **動作確認**: Revit 環境で動作確認済み
- **アイコン**: ∇FL線 + 矢印 + I型梁 + 3色ブロック（`beam_top_level_32.png`）

## RoomTagCreator（ルームタグ一括配置）設計メモ

### 概要
平面ビュー上のルームにルームタグを一括で自動配置するコマンド。

### コード構成
```
Commands/RoomTagCreator/
├── RoomTagAutoCreatorCommand.cs  # メインコマンド (IExternalCommand)
├── RoomTagService.cs             # タグ配置ロジック
├── RoomTagUI.xaml                # WPF設定ダイアログ
├── RoomTagUI.xaml.cs             # ダイアログのコードビハインド
└── Model/
    ├── RoomTagTypeInfo.cs        # タグタイプ情報モデル
    ├── LayoutSettings.cs         # レイアウト設定モデル
    └── RoomInfo.cs               # ルーム情報モデル
```

### アイコン設計
- **ファイル**: `Resources/Icons/room_tag_32.png`
- **デザイン**: 表形式アイコン（32x32 PNG、透過背景）
- **構成**: 上部1行は通し（結合セル風）、下部は3列グリッド、最下行は2列
- **色**: 黒線 `(0,0,0)` のみ、背景は透過
- **作成方法**: Python Pillow (`ImageDraw`) で生成

### アイコン作成の知見
- 32x32の小さいアイコンでは、罫線1pxでの表現が基本
- 表形式のアイコンは「通しの行」と「分割された行」の組み合わせで表現
- ユーザーのフィードバックに応じて縦罫線の有無を調整（上部セクションは通し＝縦罫線なし）
- 透過PNGで作成し、背景色は不要

### 現在のステータス
- **コード実装**: 完了（コマンド、サービス、ダイアログ、モデル）
- **アイコン**: 表形式アイコン（`room_tag_32.png`）- 上部通し行 + 下部グリッド
- **リボン登録**: 未確認（Application.cs への登録が必要）

## リボンメニュー整理

### 概要
リボンメニューを5パネル構成に整理。

### パネル構成（左から順）
1. **通り芯・レベル** — 通り芯/レベルの符号表示切替（両方/左のみ/右のみ）
2. **シート** — シート一括作成
3. **ビュー** — ビューポート位置コピー/ペースト、トリミング領域コピー/ペースト、3Dビューコピー/ペースト、セクションボックスコピー/ペースト
4. **注釈・詳細** — 部屋タグ自動配置、塗潰し領域 分割・統合
5. **構造** — 梁下端色分け、梁天端色分け、耐火被覆色分け
6. **データ** — EXCELエクスポート、EXCELインポート
7. **設定** — 言語切替（JP/US/CN）、バージョン情報、マニュアル

### 実装
- `Application.cs` の `OnStartup()` で各パネルを個別メソッドで構築
- `CreateGridLevelPanel()`, `CreateSheetPanel()`, `CreateViewPanel()`, `CreateAnnotationPanel()`, `CreateStructuralPanel()`, `CreateDataPanel()`, `CreateSettingsPanel()`

## ExcelExportImport（EXCELエクスポート/インポート）設計メモ

### 概要
Revit要素のパラメータをExcelファイルにエクスポートし、Excelで編集した値をRevitにインポートする双方向連携機能。

### コード構成
```
Commands/ExcelExportImport/
├── ExcelExportCommand.cs         # エクスポートコマンド (IExternalCommand)
├── ExcelImportCommand.cs         # インポートコマンド (IExternalCommand)
├── Models/
│   ├── CategoryInfo.cs           # カテゴリ情報モデル
│   ├── ExportSettings.cs         # エクスポート設定モデル
│   └── ParameterInfo.cs          # パラメータ情報モデル
├── Services/
│   ├── ExcelExportService.cs     # Excel書き出し処理
│   ├── ExcelImportService.cs     # Excel読み込み・比較・色付け処理
│   ├── ExcelProcessHelper.cs     # 開いているExcel検出・COM経由操作
│   ├── ParameterService.cs       # パラメータ取得/設定サービス
│   ├── RevitCategoryHelper.cs    # カテゴリ別要素取得
│   └── SettingsService.cs        # エクスポート設定の保存/読込
└── Views/
    ├── ExportDialog.xaml          # エクスポート設定ダイアログ
    ├── ExportDialog.xaml.cs       # エクスポートダイアログのコードビハインド
    ├── ImportDialog.xaml          # インポートプレビューダイアログ
    └── ImportDialog.xaml.cs       # インポートダイアログのコードビハインド
```

### 外部ライブラリ
- **ClosedXML**: Excel (.xlsx) の読み書き（NuGet パッケージ）
- **Excel COM (動的)**: 開いている Excel への直接操作（`Marshal.GetActiveObject`）

### エクスポート設計
- カテゴリ別にシートを分割（シート名=カテゴリ名）— デフォルト動作
- **シート統合モード**: ダイアログのチェックボックスで切替可能。全カテゴリを「データ」シート1枚にまとめて出力
- ヘッダー行: `要素ID | カテゴリ | パラメータ1 | パラメータ2 | ...`
- パラメータ名のプレフィックス: `I-`（インスタンス）/ `T-`（タイプ）
- ヘッダー色: 緑 `RGB(155, 187, 89)` + 白文字
- フォント: ＭＳ 明朝
- オートフィルタを自動設定

### インポート設計
- Excelファイル選択: 開いているファイル自動検出 / ファイル参照 / ドラッグ＆ドロップ
- プレビュー: 書き込み可能な変更のみ表示（読み取り専用パラメータは非表示）
- 色付け: インポート後、変更された行に背景色 `RGB(255, 255, 153)` を適用
- **セル単位の色分け**: インポート成功セルは青字 `RGB(79, 129, 189)` + 太字、失敗セルは赤字 + 太字
- **凡例**: インポート後、各シートの1行目最終列の次に `(*青字はインポート成功、赤字はインポート失敗)` を記載（青字・赤字部分は対応する色+太字）
- 色付けルート: COM経由（Excel開いている場合）→ ClosedXML（閉じている場合）
- **空セルスキップ**: Excelセルが空の場合は「変更なし」としてスキップ（シート統合モードで該当カテゴリに存在しないパラメータ列の誤検出防止）

### 開発で得た知見

#### Excel COM の `Interior.Color` 形式
- **`R + G*256 + B*65536` 形式**（VBA の `RGB()` 関数と同じ）
- **BGR ではない!** — 当初 `B + G*256 + R*65536` と誤解してRとBが逆になり、黄色のつもりが水色になった
- 例: `RGB(255, 255, 153)` = `255 + 255*256 + 153*65536` = `10092543`

#### 数値パラメータの Excel 書き込み
- `ClosedXML` の `cell.Value = stringValue` はテキスト形式で保存される → Excelで「数値が文字列として保存されています」警告が出る
- **数値は `double` 型で書き込む**: `double.TryParse` で変換してから `cell.Value = numValue`
- これによりExcelの警告が解消され、ユーザーが編集しやすくなる

#### テキスト/数値 混在時の値比較
- エクスポート時にテキストで保存 → ユーザーが編集 → 数値に変わる場合がある
- `GetString()` だけでは不十分。`cell.DataType == XLDataType.Number` をチェックし、整数なら小数点なしの文字列に変換
- 値比較は `ValuesAreEqual()` で数値比較にフォールバック（`"4700"` vs `4700.0` を同一と判定）

#### Revit パラメータの読み取り専用制限
- 構造柱の「長さ」など、Revitが自動計算するパラメータは `param.IsReadOnly = true`
- API 経由で `Set()` しても例外が発生するため、インポート時にスキップが必要
- プレビューでは読み取り専用パラメータを非表示にし、サマリーで件数と理由を表示

#### `AsValueString()` の戻り値
- `StorageType.Double` のパラメータは `AsValueString()` で表示単位での文字列を取得（例: 内部値 feet → 表示 "4700" mm）
- `SetValueString()` で表示単位の文字列からの設定が可能（内部での単位変換は自動）
- `AsValueString()` が `null` を返す場合があるため、`?? AsDouble().ToString()` でフォールバック

#### ClosedXML でのファイル読み取り（Excel 開いている場合）
- `FileShare.ReadWrite` を指定しないと、Excelがファイルをロックしているため読み取りが失敗する
- `new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)` を使用

#### エクスポート設定の保存
- ユーザーが選択したカテゴリ・パラメータの設定を JSON で保存
- 保存先: `SettingsService` でプロジェクト単位の設定ファイルを管理
- 次回エクスポート時に前回の設定を自動復元

#### ClosedXML リッチテキスト（セル内の部分書式設定）
- `cell.CreateRichText()` でリッチテキストオブジェクトを取得
- `richText.AddText("文字列")` で部分追加し、返り値に `.SetFontColor()`, `.SetBold()` で書式設定
- 1セル内で色・太字を混在させる凡例表示に使用

#### Excel COM の `Characters` による部分書式設定
- `cell.Characters[startPos, length]` で文字列の一部分を取得し、`Font.Color`, `Font.Bold` を設定可能
- **startPos は1ベース**（0ベースではない）
- 文字位置の計算は全角・半角に関わらず1文字=1カウント
- 例: `"(*青字はインポート成功、赤字はインポート失敗)"` で「青字」は位置3、「赤字」は位置14

#### シート統合モードでのインポート時の空セル問題
- シート統合モードでエクスポートすると、あるカテゴリに存在しないパラメータの列は空セルになる
- インポート時に空セル(`""`)と現在値が比較され「変更あり」と誤検出される
- **対策**: `GeneratePreview` と `Import` の両方で `string.IsNullOrEmpty(newValue)` なら処理をスキップ

#### WPF ダイアログへの Row 追加
- 既存の Grid に行を追加する場合、`<Grid.RowDefinitions>` に `<RowDefinition Height="Auto"/>` を追加
- 既存の `Grid.Row` 参照を持つ要素のインデックスも更新が必要（後続の行がずれる）

### 現在のステータス
- **機能実装**: 完了（エクスポート、インポート、プレビュー、色付け、設定保存、シート統合モード、インポート凡例）
- **動作確認**: Revit 環境で全バージョン動作確認済み
- **アイコン**: `excel_export_32.png`, `excel_import_32.png`

## 解決済み: 塗潰し領域ボタンの名称・アイコン変更がRevitに反映されない問題

### 症状
- ボタン名・アイコン・内部名を変更し、ビルド＆デプロイしても Revit に反映されない
- デバッグログも更新されない
- リボンの位置変更（パネル再構成）は反映されている

### 原因（解決済み）
**`git pull` の失敗に気づかず、古いコードでビルドしていた。**

具体的な経緯：
1. Claude Code（リモート環境）でコード変更 → コミット → push → 自動マージで main に反映
2. ユーザーのローカル（Windows）で `git pull origin main` を実行
3. **ローカルに未コミットの変更（`filled_region_32.png`）があったため `git pull` がエラーで中止された**
4. エラーメッセージを見落とし、`QuickBuild.ps1` を実行 → 古いコードのままビルド＆デプロイ
5. Revit を起動しても変更が反映されない → 様々な原因を調査（UIState.dat 削除、DLLタイムスタンプ確認等）
6. 実際にはコード自体が更新されていなかっただけだった

### 教訓・再発防止

#### Claude Code セッションでの注意事項
1. **`git pull` の結果を必ず確認させる** — エラーで中止されていないか、実際にコミットが適用されたかを確認
2. **ローカルの未コミット変更に注意** — 特にバイナリファイル（PNG等）はマージコンフリクトの原因になりやすい
3. **変更が反映されない場合、まず `git log --oneline -5` でローカルの HEAD を確認** — 期待するコミットが含まれているかが最重要
4. **デバッグログや UIState.dat 等の調査は、コードが正しくデプロイされていることを確認した後に行う**

#### 確認手順チェックリスト（変更が反映されない場合）
```
1. git status                           # 未コミット変更がないか
2. git log --oneline -5                 # 期待するコミットがHEADに含まれるか
3. git pull origin main                 # pull が成功したか（エラーなし？）
4. .\QuickBuild.ps1                     # ビルド成功を確認
5. デプロイ先DLLのタイムスタンプ確認      # 更新されているか
6. Revit 再起動して確認
```

#### git pull 失敗時の対処法
```powershell
# ローカルの未コミット変更を退避して pull
git stash
git pull origin main
git stash pop  # 必要なら退避した変更を戻す

# または、ローカル変更が不要なら破棄
git checkout -- <ファイル名>
git pull origin main
```

## 配布ZIP自動アップロード（完了）

### 概要
Revit-Add-ins でリリースすると、28yu/28tools-download の GitHub Releases にも ZIP が自動アップロードされ、配布サイトのダウンロードリンクも自動更新される。

### 自動化フロー
1. Claude Code で「v2.1をリリースして」→ `release/v2.1` ブランチを push
2. GitHub Actions が全6バージョンをビルド & ZIP作成
3. **Revit-Add-ins** の Releases に v2.1 を作成（ZIPアップロード）
4. **28tools-download** の Releases にも v2.1 を作成（ZIPアップロード）
5. 配布サイト（28tools.com）は GitHub API で latest release を自動取得 → ダウンロードリンク自動更新
6. `release/v2.1` ブランチは自動削除

### テストビルド（リリースなし）
Claude Code で「テストビルドして」→ `build/test` ブランチを push
- 全6バージョンをビルド、リリースは作成しない
- Actions の実行結果ページ → Artifacts セクションから各バージョンのZIPをダウンロード可能
- GitHub Actions 画面から手動実行も可能（「テストビルドのみ」チェックボックスをON）

### リリーストリガー方法（3種類）
| 方法 | トリガー | 用途 |
|------|---------|------|
| `release/v*` ブランチ push | Claude Code から | 通常のリリース |
| タグ push (`v*`) | ローカルPCから `git tag v2.1 && git push --tags` | 通常のリリース |
| workflow_dispatch | GitHub Actions 画面で手動実行 | 手動リリース or テストビルド |

### 設定済みの項目
- **PAT**: Classic token（`repo` スコープ）— Fine-grained token では権限不足でリリース作成が失敗する
- **Secret名**: `DOWNLOAD_SITE_TOKEN`（Revit-Add-ins リポジトリの Settings → Secrets → Actions に登録）
- **トークン指定方法**: `softprops/action-gh-release` は `with: token:` で指定（`env: GITHUB_TOKEN:` では動作しない）
- **配布サイト側**: `js/main.js` の `downloadConfig.urls` を GitHub API (`releases/latest`) で自動取得に改修済み

### リリース本文
- 両リポジトリで同じ形式（新機能テーブル + 全機能一覧）
- 新機能追加時は `build-and-release.yml` の body セクション（2箇所: Revit-Add-ins用 と 28tools-download用）を更新

### Claude Code 環境の制限
- **タグの push は 403 で失敗する**（プロキシ制限）→ `release/v*` ブランチ方式で代替
- **ブランチの push は可能** → `release/v*`, `build/*`, `claude/*` の push は正常動作
- ローカルPCからはタグ push も正常動作

### 関連情報
- **配布サイトリポジトリ**: `28yu/28tools-download`
- **配布サイトURL**: `https://28tools.com/`
- **配布方式**: GitHub Pages（リポジトリ自体がウェブサイト）
- **ZIPの配置先**: 28tools-download の GitHub Releases
- **ダウンロードパスワード**: 28tools

## FireProtection（耐火被覆範囲色分け）設計メモ

### 概要
平面・天伏・断面ビュー上の梁・柱を耐火被覆種類別に色分け塗潰しするコマンド。シートをアクティブにして実行すると、配置されている対象ビュー全てを一括処理し、自動生成した凡例（製図ビュー）をシート右上に自動配置する。

### コード構成
```
Commands/FireProtection/
├── FireProtectionCommand.cs       # メインコマンド (IExternalCommand)
├── FireProtectionDialog.xaml      # WPF設定ダイアログ
├── FireProtectionDialog.xaml.cs   # ダイアログのコードビハインド
├── FireProtectionDialogData.cs    # データモデル
├── BeamGeometryHelper.cs          # 梁ジオメトリ処理（平面/断面、柱連携クリップ）
├── FilledRegionCreator.cs         # 塗潰し領域作成 + 自動配色生成
└── LegendManager.cs               # 凡例製図ビューの作成（Excelセル表形式）
```

### 凡例シート自動配置の設計
- アクティブビューがシートで凡例ビューが作成済みの場合のみ動作
- **別トランザクション**で実行（凡例ビューをコミット後に Viewport.Create する必要があるため）
- 配置位置: 右上角固定、項目数からサイズ推定して配置
  - `estW = 85mm`（色四角列20mm + テキスト列65mm）
  - `estH = titleH + totalRows * rowH + noteH` （`noteH=45mm`は注記セクション概算）
  - 余白 `margin = 50mm` + 微調整 `upOffset = 25mm`, `rightOffset = 30mm`
- デバッグログ: `C:\temp\FireProtection_debug.txt` に VP ID と座標を記録

### 自動配色の設計
- 12色の独立パレット（赤／青／緑／黄／紫／ピンク／シアン／オレンジ／黄緑／ティール／マゼンタ／インディゴ）
- 梁用: 薄い色合い（パステル）、柱用: 濃い色合い（ビビッド）
- 13種類以上は明度を ×0.6 または +50〜80 で大きく変化させて重複回避

### 開発で得た重要な知見

#### ⚠️ アイコン PNG の DPI メタデータが必須（WPF 表示サイズ問題）
- **症状**: 自作した96px PNGをLargeImageに設定するとリボンからはみ出して表示される
- **原因**: WPF は PNG の DPI メタデータを基に論理ピクセル寸法を決定する
  - 既存の動作するアイコン（beam_top_level_96.png 等）は **DPI=288** が設定されている
  - 96px ÷ (288/96) = 32 論理ピクセル として表示される
  - DPI が未設定だと 96px がそのまま 96 論理ピクセルとして扱われ、ボタン枠を超える
- **対策**: Pillow で保存時に `dpi=(288.0106, 288.0106)` を指定する
  ```python
  img96.save(path, dpi=(288.0106, 288.0106))
  ```
- 32px版は `dpi=(96.012, 96.012)` で OK
- **検証方法**: `Image.open(path).info` で `dpi` キーがあるか確認

#### アイコンの右側カラーパレット仕様（既存と統一）
- **8×8 ピクセルの塗潰しのみ、囲い線なし**
- x範囲: 22-29、y範囲: 1-8 / 11-18 / 21-28（2px間隔）
- 色: ピンク`(255,128,148)`, 黄`(218,185,47)`, 青`(30,144,255)`
- ピクセル単位で `beam_under_level_32.png` と一致させること

#### ⚠️ トランザクションブロック内で宣言した変数のスコープ
- `using (Transaction trans) { ... }` ブロック内で宣言した変数は、ブロック終了後に使えない
- 凡例シート自動配置（別トランザクション）で `hasColumnFrame` 等のフラグを使う場合は、
  **トランザクション開始前に宣言**して、ブロック内で代入のみ行う
  ```csharp
  bool hasColumnFrame = false;
  using (Transaction trans = new Transaction(doc, "..."))
  {
      // ...
      hasColumnFrame = targetViews.Any(...);  // 代入のみ
  }
  // ここで hasColumnFrame を使える
  ```

#### Viewport.Create のサイズ取得問題
- 凡例ビューのサイズを `view.GetBoxOutline()` で取得しようとすると失敗するケースあり
  （ビューが直前のトランザクションでコミットされたばかりで、内部キャッシュが未更新）
- **対策**: 項目数からサイズを推定して直接配置する方式が確実
  ```csharp
  double estH = titleH + totalRows * rowH + noteH;
  double estW = colW + textW;
  double vpCX = outline.Max.U - margin - estW / 2.0;
  ```

#### 自動配色アルゴリズムの注意点
- ベース色を3色だけにして残りを明度シフトで生成すると、cycle≥1 のときに元色とほぼ同色になる
  （例: factor=0.95 で `(255,153,104) → (248,151,105)` で見分け不能）
- **対策**: ベースパレットを十分な数（12色程度）まで拡張し、超過時のみ明度を大きく変化させる

#### 凡例の Excel セル表形式
- タイトル行 + 各種類の色行 + 柱行 + 注記行を、表形式（横線・縦線）で構成
- 横線: 各行の境界 + 上下端
- 縦線: 左端 / 色四角列右端（colDivX=20mm）/ 右端
- 色四角は 18×8mm（行内に1mm余白）
- 注記は `※` ごとに別 TextNote、行間隔 `textHeight * 1.6 * lineCount + textHeight * 2.5`
- mm固定寸法 + テキスト垂直中央配置で安定したレイアウト

#### 柱枠線の非表示
- 柱の枠型（内側に穴の開いた矩形）の塗潰しは外周線が出るため、`SetLineStyleId()` で
  「非表示」または「Invisible」スタイルを適用して枠線を消す
  ```csharp
  var invisStyle = ... GraphicsStyle ... where Name.Contains("非表示") ...;
  if (invisStyle != null) fr.SetLineStyleId(invisStyle.Id);
  ```

### 現在のステータス
- **機能実装**: 完了（ダイアログ、梁/柱検出、自動配色、塗潰し、凡例、シート自動配置）
- **動作確認**: Revit 環境で動作確認済み
- **アイコン**: I型梁 + グレー塗潰し被覆 + 3色ブロック（`fire_protection_32.png` / `_96.png`）
- **リボン登録**: 構造パネル「色分け」セクションに追加済み

## 多言語UI（Localization）設計メモ

### 概要
リボンUI・全WPFダイアログ・TaskDialogメッセージを日本語/英語/中国語（簡体字）の3言語に対応。ランタイムで即時切替可能。

### 対応言語
| コード | 言語 | 国旗アイコン |
|--------|------|-------------|
| JP | 日本語（デフォルト） | `flag_jp_16.png` / `flag_jp_32.png` |
| US | English | `flag_us_16.png` / `flag_us_32.png` |
| CN | 简体中文 | `flag_cn_16.png` / `flag_cn_32.png` |

### アーキテクチャ

#### 多言語マネージャー (`Localization/Loc.cs`)
- 静的クラス、シングルトン的に使用
- `Loc.S("Key")` で現在の言語の文字列を取得
- フォールバック: キーが見つからない場合は日本語 → キー文字列そのもの
- 言語設定はアセンブリディレクトリの `28tools_lang.txt` に永続化
- `Loc.LanguageChanged` イベントでランタイム切替を通知

#### 文字列辞書 (`Localization/StringsJP.cs`, `StringsEN.cs`, `StringsCN.cs`)
- 各言語約325エントリの `Dictionary<string, string>`
- キー命名規則: `{カテゴリ}.{サブカテゴリ}.{項目}` 形式
  - `Ribbon.Panel.GridLevel` — パネル名
  - `Ribbon.GridBubble.Both` — ボタンテキスト
  - `Ribbon.GridBubble.Both.Tip` — ツールチップ
  - `Common.OK`, `Common.Cancel` — 共通UI文字列
  - `Sheet.Title`, `Fire.Step1.Title` — ダイアログ固有文字列

#### 言語切替コマンド (`Commands/LanguageSwitch/`)
- `SwitchToJapaneseCommand` / `SwitchToEnglishCommand` / `SwitchToChineseCommand`
- `AboutCommand` — バージョン情報ダイアログ表示
- `ManualCommand` — ブラウザでマニュアルURL (https://28tools.com/addins.html) を開く

### 設定パネル（リボン右端）の構成

3段スタック構成 (`AddStackedItems`):
1. **言語プルダウン** — `PulldownButton` + 3つのサブボタン（国旗アイコン付き）
2. **バージョン情報** — ⓘ アイコン (`ver_16.png`)、`AboutCommand` を実行
3. **マニュアル** — ? アイコン (`manual_16.png`)、`ManualCommand` を実行

### ランタイム言語切替の仕組み

1. ユーザーが言語プルダウンからJP/US/CNを選択
2. `SwitchTo*Command` → `Loc.SetLanguage()` 実行
3. `Loc.LanguageChanged` イベント発火
4. `Application.UpdateRibbonLanguage()` が呼ばれ:
   - 全パネルタイトルを更新（`_panelKeys` 辞書）
   - 全ボタンテキストを更新（`_buttonTextKeys` 辞書）
   - 全ツールチップを更新（`_buttonTipKeys` 辞書）
   - 言語プルダウンの国旗アイコンを動的に差し替え
5. WPFダイアログは次回表示時に `ApplyLocalization()` で翻訳適用

### WPFダイアログの多言語化パターン

```csharp
public DialogName(/* params */)
{
    InitializeComponent();
    ApplyLocalization();  // 翻訳を即座に適用
}

private void ApplyLocalization()
{
    this.Title = Loc.S("DialogName.Title");
    labelX.Text = Loc.S("DialogName.LabelX");
    btnOK.Content = Loc.S("Common.OK");
}
```

多言語化済みダイアログ（全8種）:
- SheetCreationDialog, BeamTopLevelDialog, BeamUnderLevelDialog
- FilledRegionSplitMergeDialog, FireProtectionDialog
- ExportDialog, ImportDialog, RoomTagUI

### コード構成
```
Commands/LanguageSwitch/
├── LanguageSwitchCommand.cs      # 3言語切替コマンド (JP/US/CN)
├── AboutCommand.cs               # バージョン情報表示
└── ManualCommand.cs              # マニュアルURLを開く

Localization/
├── Loc.cs                        # 多言語マネージャー (静的クラス)
├── StringsJP.cs                  # 日本語辞書 (~325エントリ)
├── StringsEN.cs                  # 英語辞書
└── StringsCN.cs                  # 中国語（簡体字）辞書
```

### 開発で得た知見

#### リボンボタンのランタイム更新
- `RibbonItem.ItemText` でボタンテキストを変更可能
- `PulldownButton.Image` でアイコンを動的に差し替え可能
- パネルタイトルは `RibbonPanel.Title` で変更可能
- ボタンとパネルの参照は `Application.cs` のフィールドに保持しておく必要がある
  （`_panels`, `_buttons`, `LanguagePulldown` 等）

#### ボタン名⇔ローカライゼーションキーのマッピング
- `_buttonTextKeys` / `_buttonTipKeys` で内部ボタン名とキーを対応付け
- `_panelKeys` はパネルのインデックスベースの配列
- ボタン追加時にマッピングも同時に追加しないと言語切替で更新されない

#### 国旗アイコンの動的差し替え
- `LoadImage()` で `pack://application:,,,/` URI から読み込み → `BitmapImage.Freeze()` 必須
- 言語コード → ファイル名の変換: `$"flag_{Loc.CurrentLang.ToLower()}_16.png"`
- 16px版をスタックボタンの `Image` に設定、32px版はプルダウンサブボタンの `LargeImage`

#### 翻訳辞書の管理
- 3言語ファイルのキーは完全に一致させる（キー不一致時はJPにフォールバック）
- 新機能追加時は3ファイル全てに同時にエントリを追加
- TaskDialogメッセージもダイアログ同様に `Loc.S()` で翻訳可能

#### 設定パネルの3段スタック
- `AddStackedItems()` は2個または3個の `RibbonItemData` を受け取る
- 3段スタック時はアイコンサイズ16px、テキストは短めにする
- プルダウンボタンもスタックアイテムとして配置可能

### 現在のステータス
- **機能実装**: 完了（言語切替、リボン全体更新、WPFダイアログ8種、TaskDialogメッセージ）
- **対応言語**: 日本語（デフォルト）、英語、中国語（簡体字）
- **アイコン**: 国旗3種 (`flag_jp/us/cn`)、ⓘバージョン情報 (`ver`)、?マニュアル (`manual`)
- **動作確認**: 未確認（Revit環境でのテストが必要）
