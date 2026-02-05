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

## 📞 サポート

問題が発生した場合:
1. このドキュメントのトラブルシューティングを確認
2. GitHubのIssuesで報告: https://github.com/28yu/Revit-Add-ins/issues
3. CLAUDE.mdを参照
