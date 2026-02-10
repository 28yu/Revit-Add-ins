# Tools28 - Revit Addin

Tools28は、Autodesk Revitの作業効率を向上させる便利なツール集です。

## 対応バージョン

- Revit 2022
- Revit 2023
- Revit 2024
- Revit 2025
- Revit 2026

## 機能

### 1. シート作成 (Sheet Creation)
複数のシートを効率的に作成します。

### 2. ビューコピー (View Copy)
ビューの設定をコピー＆ペーストします。

### 3. グリッドバブル (Grid Bubble)
グリッドのバブル表示を制御します。

### 4. セクションボックスコピー (Section Box Copy)
セクションボックスの設定をコピー＆ペーストします。

### 5. ビューポート位置調整 (Viewport Position)
ビューポートの位置を簡単に調整します。

### 6. クロップボックスコピー (Crop Box Copy)
クロップボックスの設定をコピー＆ペーストします。

## 開発環境のセットアップ

> **🚀 初めての方は [QUICK_START.md](QUICK_START.md) を参照してください**
> - PowerShellの使い方
> - ビルド＆デプロイの詳細手順
> - トラブルシューティング

### 必要な環境
- Windows 10/11
- Visual Studio 2019以降（MSBuildが必要）
- .NET Framework 4.8
- Revit 2022-2026のいずれか

### クローンとビルド

```bash
# リポジトリをクローン
git clone https://github.com/[your-username]/Tools28.git
cd Tools28

# すべてのバージョンをビルド
powershell -ExecutionPolicy Bypass -File .\BuildAll.ps1

# .addinファイルを生成
powershell -ExecutionPolicy Bypass -File .\GenerateAddins.ps1

# 配布パッケージを作成
powershell -ExecutionPolicy Bypass -File .\CreatePackages.ps1
```

### 単一バージョンのビルド

特定のRevitバージョンのみをビルドする場合：

```powershell
# Revit 2024用にビルド
msbuild Tools28.csproj /p:Configuration=Release /p:RevitVersion=2024
```

## プロジェクト構造

```
Tools28/
├── Commands/                    # コマンド実装
│   ├── CropBoxCopy/            # クロップボックスコピー
│   ├── GridBubble/             # グリッドバブル
│   ├── SectionBoxCopy/         # セクションボックスコピー
│   ├── SheetCreation/          # シート作成
│   ├── ViewCopy/               # ビューコピー
│   └── ViewportPosition/       # ビューポート位置調整
├── Resources/                   # リソースファイル
│   └── Icons/                  # アイコン
├── Properties/                  # アセンブリ情報
├── Application.cs              # メインアプリケーション
├── Tools28.csproj              # プロジェクトファイル
├── BuildAll.ps1                # 全バージョンビルドスクリプト
├── GenerateAddins.ps1          # .addin生成スクリプト
└── CreatePackages.ps1          # パッケージ作成スクリプト
```

## 開発ワークフロー

### 日常的な開発（推奨）

```powershell
# 1. コードを修正

# 2. クイックビルド＆デプロイ（Revit 2022向け、高速）
.\QuickBuild.ps1

# 3. Revit 2022を起動してテスト
```

### リリース前の準備

```powershell
# 1. 全バージョン（2021-2026）をビルド
.\BuildAll.ps1

# 2. 配布パッケージを作成
.\CreatePackages.ps1 -Version "1.0"

# 3. Dist\ フォルダに配布ZIPが生成される
```

### 詳細な開発ワークフロー

#### 1. コードの修正
任意のエディタまたはVisual Studioでソースコードを編集します。

#### 2. ビルド＆デプロイ

**開発時（高速）**:
```powershell
# Revit 2022向けにビルド＆デプロイ（自動）
.\QuickBuild.ps1

# 別のバージョンを指定
.\QuickBuild.ps1 -RevitVersion 2024
```

**リリース時（全バージョン）**:
```powershell
# 全バージョンビルド
.\BuildAll.ps1
```

#### 3. テスト
- **QuickBuild.ps1の場合**: 自動デプロイ済み、Revitを起動/再起動するだけ
- **BuildAll.ps1の場合**: 手動でDLLをコピーが必要

#### 4. パッケージ作成
```powershell
.\CreatePackages.ps1 -Version "1.0"
# → Dist\28Tools_Revit20XX_v1.0.zip が生成される
```

## Claude Codeでの開発

このプロジェクトは、Claude Codeを使用した開発に最適化されています。

```bash
# 機能を追加する場合の例
claude-code "新しいコマンドを追加してください：壁の高さを一括変更する機能"

# バグを修正する場合の例
claude-code "SheetCreationCommandsでダイアログが表示されない問題を修正"

# リファクタリングの例
claude-code "ViewCopyCommandsのコードをクリーンアップして可読性を向上"
```

## マルチバージョン対応の仕組み

このプロジェクトは、条件付きコンパイルを使用して単一のソースコードから複数のRevitバージョンに対応したDLLを生成します。

### バージョン固有のコードを書く場合

```csharp
#if REVIT2022
    // Revit 2022固有のコード
#elif REVIT2023
    // Revit 2023固有のコード
#elif REVIT2024
    // Revit 2024固有のコード
#else
    // その他のバージョン
#endif
```

## トラブルシューティング

### ビルドが失敗する場合

1. Revit APIのパスを確認
   - `Tools28.csproj`の`<HintPath>`が正しいか確認
   
2. MSBuildが見つからない場合
   - Visual Studio Installerで「MSBuild」がインストールされているか確認

3. 権限エラーが発生する場合
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## ライセンス

このプロジェクトは[MITライセンス](LICENSE)の下で公開されています。

## 貢献

プルリクエストを歓迎します！大きな変更の場合は、まずIssueを開いて変更内容について議論してください。

## サポート

問題や質問がある場合は、[GitHub Issues](https://github.com/[your-username]/Tools28/issues)でお知らせください。
