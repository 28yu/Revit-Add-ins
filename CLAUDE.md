# Tools28 - Revit Add-in 開発ガイド (Claude Code用)

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
├── Commands/                   # 機能コマンド群
│   ├── GridBubble/             # 通り芯・レベルの符号表示切替
│   ├── SheetCreation/          # シート一括作成 (WPFダイアログ付き)
│   ├── ViewCopy/               # 3Dビュー視点コピー
│   ├── SectionBoxCopy/         # セクションボックスコピー
│   ├── ViewportPosition/       # ビューポート位置コピー (自動マッチング)
│   └── CropBoxCopy/            # トリミング領域コピー
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
│   └── Tools28.addin            # マニフェストファイル
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
- `C:\ProgramData\Autodesk\Revit\Addins\{VERSION}\` へ DLL/addin をコピー
- ディレクトリ不在時は自動作成
- 28Tools フォルダ、DLL、addin の存在確認とエラーハンドリング
- コピー結果の表示

### uninstall.bat の内容
- `C:\ProgramData\Autodesk\Revit\Addins\{VERSION}\` から Tools28.dll / Tools28.addin を削除

### README.txt の内容
- クイックスタート手順 (install.bat を管理者実行 → Revit 再起動)
- 機能一覧
- アンインストール手順
- 対応バージョン

### インストール先
`C:\ProgramData\Autodesk\Revit\Addins\{VERSION}\`

### 配布ZIP作成手順
```powershell
# 1. 全バージョンをビルド
.\BuildAll.ps1

# 2. 配布ZIPを作成 (バージョン番号を指定)
.\CreatePackages.ps1 -Version "1.0"

# 出力先: .\Dist\28Tools_Revit20XX_v1.0.zip
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

### 2. リボンへの登録

`Application.cs` の `OnStartup()` メソッド内でボタンを追加：

```csharp
PushButton btn = panel.AddItem(new PushButtonData(
    "FeatureName",
    "機能名",
    assemblyPath,
    "Tools28.Commands.FeatureName.FeatureNameCommand"
)) as PushButton;
btn.ToolTip = "機能の説明";
```

### 3. アイコンの追加（オプション）

`Resources/Icons/` に32x32 PNGを追加し、`.csproj` に `<Resource>` を追加

```xml
<ItemGroup>
  <Resource Include="Resources\Icons\FeatureName.png" />
</ItemGroup>
```

### 4. ビルド＆テスト

```powershell
.\QuickBuild.ps1  # Revit 2022でビルド→デプロイ
```

※ SDK-style csproj のため `.cs` ファイルは自動認識される（`<Compile Include>` は不要）

## 外部参照

- **マニュアル**: https://28yu.github.io/28tools-manual/
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

### Revit API 参照
NuGet パッケージ `Nice3point.Revit.Api` を使用 (ローカル Revit 不要)

### 自動マージ＆デプロイ (`claude/**` ブランチ)

`claude/` で始まるブランチを push すると、以下が自動実行される：

1. **auto-merge-claude.yml**: main へ直接 squash merge → ブランチ削除（PR不要）
2. **deploy-pages.yml**: main への push / merge をトリガーに `Docs/` を GitHub Pages へデプロイ
3. **ローカル自動ビルド**: `AutoBuild.ps1` が main の更新を検知 → 自動 pull → ビルド → デプロイ → Windows 通知

#### ⚠️ push 前に必ず rebase すること

自動マージはコンフリクトがあると失敗する。**push 前に必ず以下を実行**：

```bash
git fetch origin main
git rebase origin/main
# コンフリクトがあれば解決
git push -u origin claude/<branch-name>
```

これを怠ると PR が作成されてもマージできず、手動でのコンフリクト解消が必要になる。
特に同じファイル（例: `Docs/workflow-diagram.html`）を連続して変更する場合は要注意。

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
- push 前に必ず `git fetch origin main && git rebase origin/main` を実行
- 自動マージ成功後はリモートブランチが自動削除される
- 削除後に再 push する場合は `--force-with-lease` ではなく通常の push を使用
- マージ失敗時は rebase して再 push すれば自動リトライされる

### 現在のステータス
- **機能実装**: 完了（ダイアログ、計算、フィルタ、ラベル、凡例）
- **アイコン**: 案A採用済み（I型梁 + カラーバー、`beam_under_level_32.png`）
- **未テスト**: 実際の Revit 環境での動作確認が必要
