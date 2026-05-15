# Tools28 - Revit Add-in 開発ガイド (Claude Code用)

> 現在の作業状況 → **STATUS.md** / TODOリスト → **TASKS.md** / 開発知見 → **Docs/DEVLOG.md**

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
├── QuickBuild.ps1              # 高速ビルド＆デプロイ（開発用）
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
   - タイトルのバージョン番号とリリースノートを更新

3. **`Packages/{VERSION}/install.bat`** (全6バージョン)
   - ヘッダーのバージョン番号を更新

4. **`.github/workflows/build-and-release.yml`**
   - Release body の「機能一覧」セクションに新機能を追加

5. **`Application.cs`**
   - リボンに新しいボタンを登録
   - `_buttonTextKeys` / `_buttonTipKeys` にマッピングを追加（言語切替時の動的更新用）

```bash
# README.txt / install.bat のバージョン番号を一括置換
for ver in 2021 2022 2023 2024 2025 2026; do
  sed -i 's/v2.0/v2.1/g' Packages/$ver/README.txt
  sed -i 's/v2.0/v2.1/g' Packages/$ver/install.bat
done
```

## 開発ワークフロー

```powershell
# 日常開発サイクル
.\QuickBuild.ps1  # Revit 2022でビルド→デプロイ → Revit起動してテスト

# リリース準備
.\BuildAll.ps1
.\CreatePackages.ps1 -Version "1.1"
git add .
git commit -m "機能追加"   # ← commit と push は別々の Bash 呼び出しで！
git push -u origin claude/branch-name
```

## 新機能追加手順

### 1. コマンドクラスの作成

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

コード内の文字列は `Loc.S()` で置換:
```csharp
btn.ToolTip = Loc.S("Ribbon.FeatureName.Tip");
TaskDialog.Show(Loc.S("Common.Error"), Loc.S("FeatureName.SelectPrompt"));
```

WPFダイアログがある場合は `ApplyLocalization()` をコンストラクタで呼び出す:
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

```csharp
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

### 4. アイコンの追加（オプション）

```xml
<ItemGroup>
  <Resource Include="Resources\Icons\FeatureName.png" />
</ItemGroup>
```

※ SDK-style csproj のため `.cs` ファイルは自動認識される（`<Compile Include>` は不要）

### 5. マニュアルの作成

`Docs/Features/FeatureName.md` を作成する（日本語）。

### 6. features.json に追加（必須）

`Docs/features.json` の `features` 配列に1エントリ追加する。**これを忘れると配布サイトのカードが追加されない。**

```json
{
  "id": "FeatureName",
  "category": "structure",
  "icon": "icons/features/feature_name.png",
  "manual": "Features/FeatureName.md",
  "added_in": "2.2",
  "names": {
    "ja": "機能名（日本語）",
    "en": "Feature Name",
    "zh": "功能名称"
  }
}
```

**`added_in` はそのリリースバージョンを入れる（例: `"2.2"`）。**  
これにより、そのバージョンのリリース本文に「⭐新機能」として自動掲載される。

**カテゴリ ID 一覧**（追加可能、変更・削除は禁止）:

| ID | 日本語名 |
|--|--|
| `grid` | 通り芯・レベル |
| `sheet` | シート |
| `view` | ビュー |
| `annotation` | 注釈・詳細 |
| `structure` | 構造 |
| `data` | データ連携 |
| `settings` | 設定 |

## 外部参照

- **マニュアル**: https://28tools.com/addins.html
- **配布サイト**: https://28yu.github.io/28tools-download/
- **リポジトリ**: https://github.com/28yu/Revit-Add-ins

## CI/CD (GitHub Actions)

### 自動リリース
- `release/v*` ブランチ push → 全6バージョンビルド → 配布ZIP作成 → GitHub Releases
- タグ push (`v*`) でも同様
- GitHub → Actions → "Build and Release" → Run workflow で手動実行可能

### 自動マージ (`claude/**` ブランチ)
- `claude/` で始まるブランチを push → main へ squash merge → ブランチ削除（PR不要）
- push 前の rebase は PreToolUse hook で自動実行（`.claude/settings.json`）
- **⚠️ commit と push は必ず別の Bash 呼び出しで実行すること**

### テストビルド
- `build/test2` ブランチを push → ビルドのみ、リリースなし（`build/test` は 403 になるため `test2` を使う）

### AutoBuild（バックグラウンド自動ビルド）
- **起動方法**: `StartAutoBuild.vbs` をダブルクリック → UAC「はい」
- **動作**: 30秒間隔で `origin/main` を監視 → 変更検知 → pull → ビルド → デプロイ → 通知
- **ファイル構成**: `StartAutoBuild.vbs`（管理者昇格）/ `AutoBuild.ps1`（監視ループ）/ `AutoBuild.log`
- **再起動**: タスクマネージャーで `powershell.exe` を終了 → `StartAutoBuild.vbs` を再実行

### ⚠️ ローカルでの pull 失敗に注意
- ローカルに未コミット変更があると `git pull` がエラーで中止される
- pull 後は `git log --oneline -3` で期待するコミットが反映されたか確認する
- 対処: `git stash` → `git pull origin main` → `git stash pop`

## GitHub Pages (Docs/)

- **URL**: `https://28yu.github.io/Revit-Add-ins/`
- **ワークフロー図**: `Docs/workflow-diagram.html`

## 注意事項

- Revit API は NuGet パッケージ経由で取得 (`Nice3point.Revit.Api.RevitAPI` / `RevitAPIUI`)
- トランザクションは `TransactionMode.Manual` を使用
- デバッグログは `C:\temp\Tools28_debug.txt` に出力
- WPFダイアログを使用するコマンドは XAML + コードビハインドで構成

## Revit API の既知の制限事項

- **ラインワークツール (Linework Tool) のオーバーライドは API で取得・設定不可** — 詳細は `DEVELOPMENT.md` 参照

## 実装済み機能

| 機能 | コード構成 | 詳細設計 |
|------|----------|---------|
| BeamUnderLevel（梁下端色分け） | `Commands/BeamUnderLevel/` | `Docs/DEVLOG.md#BeamUnderLevel` |
| BeamTopLevel（梁天端色分け） | `Commands/BeamTopLevel/` | `Docs/DEVLOG.md#BeamTopLevel` |
| RoomTagCreator（部屋タグ自動配置） | `Commands/RoomTagCreator/` | `Docs/DEVLOG.md#RoomTagCreator` |
| FilledRegionSplitMerge（塗潰し領域分割統合） | `Commands/FilledRegionSplitMerge/` | `Docs/DEVLOG.md#FilledRegionSplitMerge` |
| ExcelExportImport（Excel連携） | `Commands/ExcelExportImport/` | `Docs/DEVLOG.md#ExcelExportImport` |
| FireProtection（耐火被覆色分け） | `Commands/FireProtection/` | `Docs/DEVLOG.md#FireProtection` |
| FormworkCalculator（型枠数量算出） | `Commands/FormworkCalculator/` | `Docs/DEVLOG.md#FormworkCalculator` |
| LanguageSwitch / Localization（多言語UI） | `Commands/LanguageSwitch/` + `Localization/` | `Docs/DEVLOG.md#LocSystem` |

### リボンパネル構成（左から順）
1. **通り芯・レベル** — 符号表示切替（両方/左/右）
2. **シート** — シート一括作成
3. **ビュー** — ビューポート位置/トリミング領域/3Dビュー/セクションボックスのコピー
4. **注釈・詳細** — 部屋タグ自動配置、塗潰し領域 分割・統合
5. **構造** — 梁下端色分け、梁天端色分け、耐火被覆色分け、型枠数量算出
6. **データ** — EXCELエクスポート、EXCELインポート
7. **設定** — 言語切替（JP/US/CN）、バージョン情報、マニュアル

実装: `Application.cs` の `CreateGridLevelPanel()` 等、パネル別メソッドで構築

## 段階的な有償化戦略

| フェーズ | 内容 | 状態 |
|---------|------|------|
| Phase 1 | バージョン有効期限（`Licensing/ExpiryManager.cs`） | 実装済み (v2.1〜) |
| Phase 2 | ライセンスキー方式（`license.dat` + RSA署名） | 有償化時に実装 |
| Phase 3 | 機能の有償化階層（無料/Standard/Pro） | 有償化時に検討 |
| Phase 4 | 旧無料版ユーザーの自然移行（有効期限で解決） | Phase1依存 |

避けるべき対策: 難読化・オンライン認証サーバー・マシンロック（コスト・安定性リスクが大）

### ⚠️ 新バージョンリリース時の作業
1. `Licensing/ExpiryManager.cs` の `ExpiryDate` を **1年後** に更新
2. `Properties/AssemblyInfo.cs` のバージョン番号更新と同時に行う
3. リリースノートに有効期限を明記

## 多言語対応のよくあるバグパターン

### ❌ 最頻出バグ: ローカライゼーションキー名の不一致

**症状**: ダイアログにキー文字列がそのまま表示される（例: `Export.SelectCategory.Header`）

**原因**: `Loc.S("...")` に渡すキー名が辞書に存在しない（フォールバックでキー文字列を返す）

**修正パターン**:
```csharp
// ❌ 辞書にないキー名
grpCategory.Header = Loc.S("Export.SelectCategory.Header");
// ✅ 辞書にあるキー名
grpCategory.Header = Loc.S("Export.Category");
```

**チェック方法**: `ApplyLocalization()` 内の全 `Loc.S("...")` キーが `StringsJP.cs` に存在するか確認
```bash
grep -o '"[^"]*"' YourDialog.xaml.cs | grep Loc.S | while read key; do grep $key Localization/StringsJP.cs; done
```

過去事例一覧: `Docs/DEVLOG.md#多言語バグ過去事例` 参照

### カテゴリ名ローカライズ（ExcelExportImport）
- `CategoryInfo.Name`: Revit提供のカテゴリ名（英語）— 設定保存・内部使用。変更しないこと
- `CategoryInfo.DisplayLabel`: UI表示用（`CategoryLocalizer.GetLocalizedName()` 経由）
- `CategoryLocalizer`: BuiltInCategory → `Category.*` Loc.Sキーのマッピング（約60カテゴリ）
