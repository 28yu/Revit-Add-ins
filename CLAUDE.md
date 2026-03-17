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
│   ├── CropBoxCopy/            # トリミング領域コピー
│   └── BeamTopLevel/           # 梁天端レベル色分け (平面ビュー)
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
- push 前に必ず `git fetch origin main && git rebase origin/main` を実行
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
5. **構造** — 梁下端色分け、梁天端色分け

### 実装
- `Application.cs` の `OnStartup()` で各パネルを個別メソッドで構築
- `CreateGridLevelPanel()`, `CreateSheetPanel()`, `CreateViewPanel()`, `CreateAnnotationPanel()`, `CreateStructuralPanel()`

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
