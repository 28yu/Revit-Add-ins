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
├── .github/workflows/          # GitHub Actions (自動ビルド・リリース)
│   └── build-and-release.yml   #   タグ push or 手動実行で配布ZIP生成
├── Dist/                       # 配布ZIP出力先 (git管理外)
├── QuickBuild.ps1              # 🚀 高速ビルド＆デプロイ（開発用）
├── BuildAll.ps1                # 全バージョン一括ビルド（リリース用）
├── GenerateAddins.ps1          # .addinマニフェスト生成
├── CreatePackages.ps1          # 配布ZIP作成
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

## 注意事項

- Revit API は NuGet パッケージ経由で取得 (`Nice3point.Revit.Api.RevitAPI` / `RevitAPIUI`)
- トランザクションは `TransactionMode.Manual` を使用
- デバッグログは `C:\temp\Tools28_debug.txt` に出力
- WPFダイアログを使用するコマンドは XAML + コードビハインドで構成

---

## 🚨 重要：FilledRegion機能の現状（2026-02-10）

### 現状
**FilledRegion機能（塗り潰し領域 分割/統合）は一時的に無効化されています。**

理由：Revit 2022でのビルドエラーを回避するため

### 無効化した箇所
1. **Application.cs**: `CreateDetailPanel()` メソッド内のFilledRegionボタン登録をコメントアウト
2. **Tools28.csproj**:
   - FilledRegionSplitMergeDialog.xaml をコメントアウト
   - `Commands\FilledRegionSplitMerge\**\*.cs` を除外

### ビルドエラーの詳細

#### エラー内容
```
error CS1061: 'Document' に 'NewFilledRegion' の定義が含まれておらず、
型 'Document' の最初の引数を受け付けるアクセス可能な拡張メソッド 'NewFilledRegion' が見つかりませんでした。
```

#### エラー発生場所
- `Commands\FilledRegionSplitMerge\FilledRegionHelper.cs` (123行目, 165行目)

#### 調査結果
- **コードは正しい**: `doc.Create.NewFilledRegion(view, typeId, loops);` を使用
- **using文も正しい**: `using Autodesk.Revit.DB;` が存在
- **Revit 2022特有の問題**: Revit 2021では動作する可能性あり
- **NuGetパッケージ**: Nice3point.Revit.Api.RevitAPI バージョン 2022.* を使用

### 影響範囲
- ✅ **他の全機能は正常に動作**（シート作成、ビューコピー、グリッドバブル等）
- ✅ **アイコン更新は完了**（`filled_region_32.png` は最新版）
- ❌ **FilledRegion機能のみ使用不可**（「詳細」パネルに「領域」ボタンが表示されない）

### 次のステップ（未対応）

#### プランB：Revit 2022での問題を詳細調査
1. **Revit 2022 API仕様の確認**
   - `Document.Create.NewFilledRegion` メソッドの署名を確認
   - パラメータの型や順序が正しいか検証
   - Revit 2022で変更があったか調査

2. **段階的なデバッグ**
   - 最小限のコードでテスト
   - using文の再確認
   - NuGetパッケージのバージョン確認

3. **代替APIの検討**
   - Revit 2022で推奨される別の方法があるか調査
   - 条件付きコンパイルで2022専用コードを作成

4. **コードの修正と検証**
   - 必要に応じてFilledRegionHelper.csを修正
   - Revit 2022でビルド＆テスト
   - Revit 2021でも動作確認（後方互換性）

### 再有効化手順（問題解決後）

1. **Tools28.csprojの修正**
   ```xml
   <!-- コメントアウトを解除 -->
   <Page Include="Commands\FilledRegionSplitMerge\FilledRegionSplitMergeDialog.xaml">

   <!-- 除外を削除 -->
   <Compile Remove="Commands\FilledRegionSplitMerge\**\*.cs" />  <!-- この行を削除 -->
   ```

2. **Application.csの修正**
   ```csharp
   // CreateDetailPanel() メソッド内のコメントアウトを解除
   PushButtonData filledRegionButtonData = new PushButtonData(...);
   ```

3. **ビルド＆デプロイ**
   ```powershell
   .\QuickBuild.ps1
   ```

4. **Revit 2022で動作確認**

### 関連ファイル
- `Commands/FilledRegionSplitMerge/FilledRegionHelper.cs` - 主要ロジック
- `Commands/FilledRegionSplitMerge/FilledRegionSplitMergeCommand.cs` - コマンドエントリポイント
- `Commands/FilledRegionSplitMerge/FilledRegionSplitMergeDialog.xaml(.cs)` - WPFダイアログ
- `Docs/Features/FilledRegionSplitMerge.md` - 機能仕様書

### アイコン更新作業（完了）✅

#### 実施内容
- `Tools/IconGenerator.html` を作成（アイコン生成ツール）
- `filled_region_32.png` を更新（4つの小さな正方形、2×2配置、ハッチング付き）
- squareSize: 8px (4px × 2)、lineWidth: 1.0px、hatchWidth: 0.6px
- `Tools/README.md` を作成（アイコン生成手順）
- `QUICK_START.md` を作成（ビルド＆デプロイの詳細ガイド）

#### 成果
- 全機能のアイコンが正常に表示
- ビルドが成功（FilledRegion機能を除く）
- Revit 2022で動作確認済み

