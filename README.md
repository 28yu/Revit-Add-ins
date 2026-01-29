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

### 1. コードの修正
任意のエディタまたはVisual Studioでソースコードを編集します。

### 2. ビルド
```powershell
.\BuildAll.ps1
```

### 3. テスト
テストしたいRevitバージョンで以下を実行：
```powershell
# DLLをRevitのアドインフォルダにコピー
Copy-Item ".\bin\Release\Revit2024\Tools28.dll" `
          "$env:APPDATA\Autodesk\Revit\Addins\2024\"
```

### 4. パッケージ作成
```powershell
.\GenerateAddins.ps1
.\CreatePackages.ps1
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
