# Tools28 - 開発者ガイド

このドキュメントは、Tools28の開発を行う開発者向けのガイドです。

## 🚀 クイックスタート

### 必要な環境

- **Visual Studio 2022** (またはそれ以降)
  - .NET デスクトップ開発ワークロード
  - .NET Framework 4.8 開発ツール
  - .NET 8.0 SDK
- **PowerShell 5.0+** (Windows標準)
- **Revit 2022** (または開発対象のバージョン)

### 初回セットアップ

```powershell
# 1. リポジトリをクローン
git clone https://github.com/28yu/Revit-Add-ins.git
cd Revit-Add-ins

# 2. 開発ブランチに切り替え
git checkout claude/setup-addon-workflow-yO1Uz

# 3. 開発バージョンを設定（dev-config.json）
# デフォルトはRevit 2022
# 他のバージョンを使う場合は dev-config.json を編集

# 4. 初回ビルド＆デプロイ
.\QuickBuild.ps1

# 5. Revit 2022を起動
# リボンに「28 Tools」タブが表示されることを確認
```

---

## 📐 新機能の設計

新機能を実装する前に、設計書を作成することを推奨します。

### 設計書テンプレートの使用

```bash
# 1. テンプレートをコピー
cp .github/FEATURE_TEMPLATE.md Docs/Features/YourFeature.md

# 2. 設計書を記入
# - 基本情報（機能名、クラス名）
# - UI設計（リボンパネル、ボタン名、アイコン）
# - 技術仕様（対象要素、API）
# - 処理フロー
# - テストケース

# 3. サンプルを参照
# Docs/Features/WallHeight-Example.md に完成したサンプルあり
```

詳細は `Docs/README.md` を参照してください。

---

## 📝 日常的な開発フロー

### 基本サイクル

```
コード修正 → QuickBuild.ps1 → Revitでテスト → 問題があれば修正
    ↑                                                ↓
    └────────────────────────────────────────────────┘
```

### 詳細ステップ

#### 1. 新機能の実装

```
Commands/配下に新しいフォルダを作成
例: Commands/WallHeight/WallHeightCommand.cs
```

```csharp
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.WallHeight
{
    [Transaction(TransactionMode.Manual)]
    public class WallHeightCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // あなたの処理をここに実装

            return Result.Succeeded;
        }
    }
}
```

#### 2. リボンに登録

`Application.cs` を開き、`OnStartup()` メソッド内で新しいボタンを追加：

```csharp
// 例: 「編集」パネルにボタンを追加
PushButton wallHeightBtn = editPanel.AddItem(new PushButtonData(
    "WallHeight",
    "壁高さ変更",
    assemblyPath,
    "Tools28.Commands.WallHeight.WallHeightCommand"
)) as PushButton;
wallHeightBtn.ToolTip = "壁の高さを一括変更";
```

#### 3. アイコンの追加（オプション）

```powershell
# 32x32 PNGアイコンを作成
# Resources/Icons/WallHeight.png

# Tools28.csproj に追加（SDK-styleなので自動認識されますが、Resourceタグが必要）
```

`Tools28.csproj` を開き、既存の `<ItemGroup>` に追加：

```xml
<ItemGroup>
  <Resource Include="Resources\Icons\WallHeight.png" />
</ItemGroup>
```

`Application.cs` でアイコンを設定：

```csharp
wallHeightBtn.LargeImage = LoadImage("WallHeight.png");
```

#### 4. ビルド＆デプロイ

```powershell
.\QuickBuild.ps1
```

**実行される処理:**
1. Revit 2022用にビルド（約10-30秒）
2. `C:\ProgramData\Autodesk\Revit\Addins\2022\` へ自動デプロイ
3. 既存のDLLは自動バックアップ

#### 5. Revitでテスト

```
1. Revit 2022を起動（または再起動）
2. 「28 Tools」タブを開く
3. 追加したボタンをクリック
4. 動作確認
```

#### 6. 問題があれば修正

```
エラーが出た場合:
- Visual Studioでデバッグ（Revitにアタッチ）
- C:\temp\Tools28_debug.txt にログ出力を追加
```

```csharp
// デバッグログの例
System.IO.File.AppendAllText(
    @"C:\temp\Tools28_debug.txt",
    $"[{DateTime.Now}] 処理開始\n"
);
```

---

## 🎯 開発バージョンの変更

異なるRevitバージョンで開発したい場合：

### 方法1: dev-config.json を編集

```json
{
  "defaultRevitVersion": "2024",
  "description": "開発時に主に使用するRevitバージョン"
}
```

その後：

```powershell
.\QuickBuild.ps1  # 2024でビルド＆デプロイ
```

### 方法2: コマンドラインで指定

```powershell
.\QuickBuild.ps1 -RevitVersion 2024
```

---

## 🏗️ プロジェクト構造

```
Revit-Add-ins/
├── Application.cs              # メインアプリ（リボンUI構築）
├── Tools28.csproj              # プロジェクトファイル
├── dev-config.json             # 開発設定（新規）
│
├── Commands/                   # 機能コマンド群
│   ├── GridBubble/
│   ├── SheetCreation/
│   ├── ViewCopy/
│   ├── SectionBoxCopy/
│   ├── ViewportPosition/
│   └── CropBoxCopy/
│
├── Resources/Icons/            # 32x32アイコン
│
├── QuickBuild.ps1              # 高速ビルド＆デプロイ（新規）
├── BuildAll.ps1                # 全バージョンビルド
├── CreatePackages.ps1          # 配布ZIP作成
└── Deploy-For-Testing.ps1      # 手動デプロイ
```

---

## 🧪 デバッグ方法

### Visual Studioでデバッグ

1. Visual Studioで `Tools28.csproj` を開く
2. デバッグ > プロセスにアタッチ
3. `Revit.exe` を選択
4. ブレークポイントを設定
5. Revitでコマンドを実行

### ログ出力

```csharp
// C:\temp\Tools28_debug.txt に出力
string logPath = @"C:\temp\Tools28_debug.txt";
System.IO.File.AppendAllText(logPath, $"[{DateTime.Now}] メッセージ\n");
```

---

## 🤖 自動ビルド & デプロイ（セルフホストランナー）

Claude Code で開発指示を出すだけで、ローカル PC 上で自動ビルド & デプロイが行われる仕組みです。

### ワークフロー

```
あなた: 開発指示 → Claude: コード編集 → push
  → GitHub Actions (auto-merge): main へ squash merge
  → GitHub Actions (local-deploy): ローカルPCでビルド & デプロイ（全自動）
  → Windows 通知「ビルド完了」
  → あなた: Revitを再起動してテスト
```

### 初回セットアップ（10分程度）

#### 1. GitHub でランナートークンを取得

1. GitHub で `28yu/Revit-Add-ins` リポジトリを開く
2. **Settings** → **Actions** → **Runners** → **New self-hosted runner**
3. OS: **Windows**、Architecture: **x64** を選択
4. 表示されるトークン（`--token XXXXX`の部分）をコピー

#### 2. Windows PC でランナーをインストール

PowerShell を**管理者として実行**し、以下を実行：

```powershell
# 1. インストール先ディレクトリを作成
mkdir C:\actions-runner
cd C:\actions-runner

# 2. ランナーをダウンロード（GitHub の画面に表示される最新版URLを使用）
Invoke-WebRequest -Uri https://github.com/actions/runner/releases/download/v2.322.0/actions-runner-win-x64-2.322.0.zip -OutFile actions-runner.zip

# 3. 解凍
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\actions-runner.zip", "$PWD")

# 4. ランナーを登録（トークンは GitHub の画面からコピー）
.\config.cmd --url https://github.com/28yu/Revit-Add-ins --token YOUR_TOKEN_HERE

# 5. Windows サービスとして登録（PC 起動時に自動実行）
.\svc.cmd install
.\svc.cmd start
```

#### 3. 動作確認

```powershell
# サービスの状態を確認
.\svc.cmd status

# GitHub の Settings → Actions → Runners で「Idle」と表示されればOK
```

### 使い方（セットアップ後）

セットアップ完了後は、以下の流れが**完全自動**です：

1. Claude Code で開発指示を出す
2. Claude がコード編集 → push
3. auto-merge → main へ反映
4. **ローカル PC で自動ビルド & デプロイ**
5. Windows デスクトップ通知が表示される
6. Revit を再起動してテスト

### 手動実行

GitHub → Actions → **Local Build & Deploy** → **Run workflow** で手動実行も可能。

### 注意事項

- **Revit 起動中の場合**: DLL がロックされてデプロイが失敗します。先に Revit を閉じてから手動で再実行してください。
- **ランナーのバージョン**: GitHub の画面に表示される最新版を使用してください（上記URLはv2.322.0の例）。
- **セキュリティ**: セルフホストランナーはプライベートリポジトリでの使用が推奨されます。パブリックリポジトリの場合、悪意ある PR からのワークフロー実行に注意が必要です。

---

## 🚢 リリース準備

開発が完了し、リリースする場合：

### 1. 全バージョンのビルド

```powershell
.\BuildAll.ps1
```

**出力先:**
```
bin\Release\Revit2021\Tools28.dll
bin\Release\Revit2022\Tools28.dll
bin\Release\Revit2023\Tools28.dll
bin\Release\Revit2024\Tools28.dll
bin\Release\Revit2025\Tools28.dll
bin\Release\Revit2026\Tools28.dll
```

### 2. 配布パッケージ作成

```powershell
.\CreatePackages.ps1 -Version "1.1"
```

**出力先:**
```
Dist\28Tools_Revit2021_v1.1.zip
Dist\28Tools_Revit2022_v1.1.zip
Dist\28Tools_Revit2023_v1.1.zip
Dist\28Tools_Revit2024_v1.1.zip
Dist\28Tools_Revit2025_v1.1.zip
Dist\28Tools_Revit2026_v1.1.zip
```

### 3. コミット＆プッシュ

```powershell
git add .
git commit -m "Add new feature: WallHeight command"
git push -u origin claude/setup-addon-workflow-yO1Uz
```

### 4. GitHub Releasesで公開

```powershell
# タグを作成してpush
git tag v1.1
git push --tags
```

**GitHub Actionsが自動実行:**
- 全6バージョンをビルド
- 配布ZIPを作成
- GitHub Releasesにアップロード

---

## 📚 参考リソース

- **Revit API ドキュメント**: https://www.revitapidocs.com/
- **RevitLookup**: デバッグ用ツール（要インストール）
- **プロジェクトREADME**: [CLAUDE.md](./CLAUDE.md)

---

## 🛠️ トラブルシューティング

### ビルドエラー: MSBuildが見つからない

```
解決策:
- Visual Studio 2022をインストール
- .NET デスクトップ開発ワークロードを有効化
```

### デプロイエラー: ターゲットディレクトリが見つからない

```
解決策:
- Revit 2022がインストールされているか確認
- C:\ProgramData\Autodesk\Revit\Addins\2022\ が存在するか確認
```

### Revitでアドインが表示されない

```
解決策:
1. Revitを完全に終了
2. タスクマネージャーでRevit.exeが終了していることを確認
3. 再度Revitを起動
4. それでもダメな場合:
   - C:\ProgramData\Autodesk\Revit\Addins\2022\Tools28.addin を確認
   - Tools28.dll が同じフォルダにあるか確認
```

### エラー: "Could not load file or assembly"

```
解決策:
- ビルドターゲットが正しいか確認（net48 or net8.0-windows）
- Nice3point.Revit.Api パッケージのバージョンを確認
- bin フォルダを削除して再ビルド
```

---

## ⚠️ Revit API の既知の制限事項

### ラインワークツールのオーバーライドは API で取得不可（2026年2月確認）

**結論**: Revit の「ラインワーク」ツール（Linework Tool）で変更されたエッジの線種を、API で検出することは**現時点では不可能**。

**背景**: ラインワークツールは、ビュー内の個別のエッジの表示スタイルを変更する機能。Revit の表示優先度で最上位（#1）に位置する。この変更を API で検出するために8つの異なるアプローチを試したが、すべて失敗した。

#### 試したアプローチと結果

| # | アプローチ | 結果 | 原因 |
|---|---|---|---|
| 1 | `Edge.GraphicsStyleId` 比較 | ✗ | ラインワーク変更を反映しない |
| 2 | 非ビュージオメトリ (`Options.View = null`) との集合比較 | ✗ | 除外範囲が広すぎる（壁の層境界線と重複） |
| 3 | Per-Edge Reference マッチング | ✗ | API がアクティブビューのデータをリーク |
| 4 | カテゴリスタイル比較 | ✗ | 塗り潰し領域を誤検出 |
| 5 | `View.Duplicate(Duplicate)` 比較 | ✗ | ラインワーク変更はモデルグラフィクス扱いで複製される |
| 6 | `ViewPlan.Create()` + エッジ比較 | ✗ | `GraphicsStyleId` が変更を反映しない |
| 7 | `FilteredElementCollector.OwnedByView()` スキャン | ✗ | ラインワークデータはビュー所有要素に存在しない |
| 8 | `CustomExporter` + `IExportContext2D` 比較 | ✗ | 2Dエクスポートデータも同一 |

#### 技術的な詳細

- `Edge.GraphicsStyleId` は要素のベースカテゴリスタイルを返す（例: 壁のエッジは常に「壁」カテゴリのスタイル）。ラインワークで `<隠線>` や `<細線>` に変更しても、API からは元のスタイルが返される。
- `IExportContext2D` はレンダリングパイプラインを通じてジオメトリを出力するが、要素ごとのエッジ数・カーブ数・線分数はオリジナルビューとクリーンビューで完全一致しており、ラインワークの影響は反映されない。
- ビュー所有要素（`OwnedByView`）のスキャンでは、FilledRegion のスケッチ線分、寸法、SketchPlane のみが見つかり、ラインワークに関連する内部要素は存在しなかった。

#### Autodesk の公式見解

Autodesk Ideas Board で公式にこの API ギャップが認められている：

> *"It looks like there is a gap here in the API. We didn't expect it to be an important workflow for API users - setting line-weights feels like a very graphical thing... but it looks like it is needed after all! We will review."*

参考リンク:
- [API access to Linework tool (Autodesk Ideas)](https://forums.autodesk.com/t5/revit-ideas/api-access-to-linework-tool-get-and-set-edge-line-overrides/idi-p/12618606)
- [API Access to Linework Tool? (Revit API Forum)](https://forums.autodesk.com/t5/revit-api-forum/api-access-to-linework-tool/td-p/6939839)

#### 今後の可能性

将来の Revit API バージョンでこの機能が追加される可能性がある。Revit 2026 API で `ParameterTypeId.EdgeLinework` プロパティが追加されているが、これは鉄筋の詳細図用であり、ラインワークツールとは無関係。新しい Revit バージョンのリリースノートを確認すること。

#### 教訓

1. **Revit API には「表示専用」の機能が存在する** — ラインワーク、一部のグラフィックオーバーライドなど、画面上の表示にのみ影響し、API からはアクセスできない機能がある。
2. **`Edge.GraphicsStyleId` は信頼できない** — ビュー固有のオーバーライドは反映されず、要素のベーススタイルのみを返す。
3. **`IExportContext2D` も万能ではない** — レンダリングパイプライン経由でも、ラインワーク情報は取得できなかった。
4. **Autodesk Ideas Board を事前に確認すること** — 実装前に API の制限事項を確認することで、無駄な開発時間を避けられる。
5. **`View.Duplicate(Duplicate)` はラインワークを保持する** — これはモデルグラフィクス扱い（アノテーションではない）のため。`ViewDuplicateOption.AsDependent` でも同様。

---

## 📞 サポート

問題が発生した場合:
1. このドキュメントのトラブルシューティングを確認
2. GitHubのIssuesで報告: https://github.com/28yu/Revit-Add-ins/issues
3. CLAUDE.mdを参照
