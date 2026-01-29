# Tools28 開発手順書

このドキュメントでは、Tools28の開発環境のセットアップから配布までの詳細な手順を説明します。

## 目次

1. [初回セットアップ](#初回セットアップ)
2. [開発ワークフロー](#開発ワークフロー)
3. [Claude Codeでの開発](#claude-codeでの開発)
4. [ビルドとテスト](#ビルドとテスト)
5. [配布パッケージの作成](#配布パッケージの作成)
6. [複数PC間での同期](#複数pc間での同期)

---

## 初回セットアップ

### 1. GitHubリポジトリの作成

```bash
# ローカルでGitリポジトリを初期化
cd Tools28
git init
git add .
git commit -m "Initial commit: Multi-version Revit addin"

# GitHubにプッシュ（事前にGitHubでリポジトリを作成）
git remote add origin https://github.com/[your-username]/Tools28.git
git branch -M main
git push -u origin main
```

### 2. 必要なソフトウェアの確認

以下がインストールされていることを確認：

- [ ] Visual Studio 2019以降（MSBuild）
- [ ] .NET Framework 4.8 SDK
- [ ] Git
- [ ] PowerShell 5.1以降
- [ ] Claude Code CLI（オプション）

### 3. Revit APIの準備

各バージョンのRevitがインストールされていない場合、以下のいずれかを実施：

**オプションA: Revitをインストール**
- 必要なバージョンのRevitをインストール

**オプションB: API DLLのみを入手**
- RevitAPI.dll と RevitAPIUI.dll を別のPCからコピー
- `Tools28.csproj`の`<HintPath>`を適切に修正

---

## 開発ワークフロー

### 基本的な開発サイクル

```
コード編集 → ビルド → テスト → コミット → プッシュ
```

### 1. コードの編集

任意のエディタでコードを編集：

- Visual Studio
- Visual Studio Code
- Rider
- Claude Code（推奨）

### 2. ビルド

**すべてのバージョンをビルド:**
```powershell
.\BuildAll.ps1
```

**特定のバージョンのみビルド:**
```powershell
# Revit 2024のみ
msbuild Tools28.csproj /p:Configuration=Release /p:RevitVersion=2024
```

### 3. テスト

#### 手動テスト手順

1. ビルドしたDLLをRevitのアドインフォルダにコピー

```powershell
# 例: Revit 2024でテストする場合
$revitVersion = "2024"
$sourceDll = ".\bin\Release\Revit$revitVersion\Tools28.dll"
$targetDir = "$env:APPDATA\Autodesk\Revit\Addins\$revitVersion\"

# .addinファイルも必要な場合は先に生成
.\GenerateAddins.ps1

Copy-Item $sourceDll $targetDir -Force
Copy-Item ".\Addins\Tools28_Revit$revitVersion.addin" `
          "$targetDir\Tools28.addin" -Force
```

2. Revitを起動してテスト

3. 問題があればログを確認
   - Revitのジャーナルファイル
   - Windowsイベントビューアー

### 4. コミットとプッシュ

```bash
git add .
git commit -m "機能: 新しいコマンドを追加"
git push
```

---

## Claude Codeでの開発

### Claude Codeのセットアップ

```bash
# Claude Code CLIをインストール（未インストールの場合）
# https://claude.ai/code から最新版をダウンロード

# プロジェクトディレクトリに移動
cd Tools28

# Claude Codeを起動
claude-code
```

### Claude Codeでの開発例

#### 新機能の追加

```
プロンプト例:
「Commands/フォルダに新しいコマンド『WallHeightAdjuster』を追加してください。
このコマンドは、選択した壁の高さを指定した値に一括変更する機能です。
UIはWPFダイアログで、高さの入力フィールドとOK/Cancelボタンを含めてください。」
```

Claude Codeが自動的に：
1. 新しいフォルダとファイルを作成
2. コマンドクラスを実装
3. WPF UIを作成
4. Application.csに登録コードを追加
5. プロジェクトファイルを更新

#### バグ修正

```
プロンプト例:
「SheetCreationCommandsで、ダイアログを閉じた後にRevitがフリーズする問題があります。
原因を調査して修正してください。」
```

#### リファクタリング

```
プロンプト例:
「ViewCopyCommandsのコードをリファクタリングしてください：
- 重複コードを削除
- メソッドを適切に分割
- コメントを追加
- 変数名を分かりやすく変更」
```

### Claude Codeでのビルド

Claude Codeはビルドも自動実行できます：

```
プロンプト例:
「すべてのRevitバージョン向けにビルドして、エラーがあれば修正してください。」
```

---

## ビルドとテスト

### ビルドスクリプトの使い方

#### BuildAll.ps1 - 全バージョンビルド

```powershell
# 基本的な使い方
.\BuildAll.ps1

# 実行時の出力例:
========================================
Tools28 - Multi-Version Build Script
========================================

MSBuild: C:\Program Files\Microsoft Visual Studio\...\MSBuild.exe

----------------------------------------
Revit 2022 をビルド中...
----------------------------------------
✓ Revit 2022 のビルドに成功しました

----------------------------------------
Revit 2023 をビルド中...
----------------------------------------
✓ Revit 2023 のビルドに成功しました

... (以下同様)

========================================
ビルド結果
========================================
Revit 2022: 成功
Revit 2023: 成功
Revit 2024: 成功
Revit 2025: 成功
Revit 2026: 成功

成功: 5 / 失敗: 0
```

### テスト手順

#### 自動デプロイスクリプトの作成（オプション）

テストを効率化するため、自動デプロイスクリプトを作成できます：

```powershell
# Deploy-For-Testing.ps1
param(
    [Parameter(Mandatory=$true)]
    [string]$RevitVersion
)

$sourceDll = ".\bin\Release\Revit$RevitVersion\Tools28.dll"
$targetDir = "$env:APPDATA\Autodesk\Revit\Addins\$RevitVersion\"

if (-not (Test-Path $sourceDll)) {
    Write-Host "エラー: DLLが見つかりません。先にビルドしてください。" -ForegroundColor Red
    exit 1
}

Copy-Item $sourceDll $targetDir -Force
Write-Host "✓ Revit $RevitVersion にデプロイしました" -ForegroundColor Green
```

使用方法：
```powershell
.\Deploy-For-Testing.ps1 -RevitVersion 2024
```

---

## 配布パッケージの作成

### 手順

#### 1. .addinファイルの生成

```powershell
.\GenerateAddins.ps1
```

これにより、`Addins/`フォルダに各バージョン用の.addinファイルが作成されます。

#### 2. 配布パッケージの作成

```powershell
.\CreatePackages.ps1
```

これにより、`Packages/`フォルダに以下が作成されます：

```
Packages/
├── Tools28_Revit2022/
│   ├── Tools28.dll
│   ├── Tools28.addin
│   └── README.txt
├── Tools28_Revit2022.zip
├── Tools28_Revit2023.zip
├── Tools28_Revit2024.zip
├── Tools28_Revit2025.zip
└── Tools28_Revit2026.zip
```

#### 3. 配布

ZIPファイルを：
- GitHubのReleasesページにアップロード
- 自社の配布サイトにアップロード
- ユーザーに直接送信

---

## 複数PC間での同期

### GitHubを使用した同期

#### 職場PCでの作業

```bash
# 作業開始時：最新の変更を取得
git pull

# コード編集

# 変更をコミット
git add .
git commit -m "機能追加: XXX"

# リモートにプッシュ
git push
```

#### 自宅PCでの作業

```bash
# 初回のみ：クローン
git clone https://github.com/[your-username]/Tools28.git
cd Tools28

# 2回目以降：最新の変更を取得
git pull

# 作業を続ける...
```

### ブランチ戦略（推奨）

```bash
# 新機能開発用のブランチを作成
git checkout -b feature/wall-height-adjuster

# 開発

# コミット
git add .
git commit -m "進捗: 壁高さ調整機能の実装"

# プッシュ
git push -u origin feature/wall-height-adjuster

# 開発完了後、メインブランチにマージ
git checkout main
git merge feature/wall-height-adjuster
git push
```

### VSCodeとGitの統合（推奨）

Visual Studio Codeを使用すると、GUIでGit操作が簡単に：

1. VSCodeでフォルダを開く
2. 左側の「ソース管理」アイコンをクリック
3. 変更をステージング
4. コミットメッセージを入力
5. 「同期」ボタンをクリック

---

## トラブルシューティング

### ビルドエラー

#### エラー: RevitAPI.dllが見つからない

**原因:** Revitがインストールされていない、またはパスが間違っている

**解決策:**
```xml
<!-- Tools28.csprojの該当箇所を修正 -->
<Reference Include="RevitAPI">
  <HintPath>C:\path\to\your\RevitAPI.dll</HintPath>
  <Private>False</Private>
</Reference>
```

#### エラー: MSBuildが見つからない

**解決策:**
Visual Studio Installerで「.NETデスクトップ開発」ワークロードをインストール

### Git関連

#### エラー: Permission denied (publickey)

**解決策:**
SSH鍵を設定するか、HTTPSでクローン：
```bash
git clone https://github.com/[your-username]/Tools28.git
```

#### コンフリクトが発生した場合

```bash
# 現在の変更を一時保存
git stash

# 最新を取得
git pull

# 変更を再適用
git stash pop

# コンフリクトを手動で解決後
git add .
git commit -m "マージコンフリクトを解決"
```

### Revitでアドインが読み込まれない

1. .addinファイルの場所を確認
   - `%APPDATA%\Autodesk\Revit\Addins\[version]\`

2. .addinファイルの内容を確認
   - DLLのパスが正しいか
   - AddInIdが正しいか

3. Revitのジャーナルファイルを確認
   - エラーメッセージを確認

4. DLLの依存関係を確認
   ```powershell
   # Dependency Walkerなどのツールで確認
   ```

---

## 参考リンク

- [Revit API ドキュメント](https://www.revitapidocs.com/)
- [Git ドキュメント](https://git-scm.com/doc)
- [Claude Code ドキュメント](https://claude.ai/code/docs)
- [PowerShell ドキュメント](https://docs.microsoft.com/powershell/)

---

## よくある質問（FAQ）

### Q: 新しいRevitバージョンが出たらどうする？

A: 以下の手順で対応：
1. `Tools28.csproj`に新しいバージョンの参照を追加
2. `BuildAll.ps1`、`GenerateAddins.ps1`、`CreatePackages.ps1`に新バージョンを追加
3. テストしてビルド

### Q: 既存のコマンドを削除したい

A: 以下のファイルを編集：
1. 該当のCommandsフォルダを削除
2. `Tools28.csproj`から`<Compile>`エントリを削除
3. `Application.cs`から登録コードを削除

### Q: アイコンを変更したい

A: `Resources/Icons/`フォルダ内の.pngファイルを置き換え

### Q: Claude Codeなしで開発できる？

A: はい、Visual StudioやVS Codeなど、任意のエディタで開発可能です。

---

以上が開発手順書です。質問がある場合は、GitHubのIssuesでお知らせください。
