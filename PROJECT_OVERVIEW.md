# Tools28 - マルチバージョン対応プロジェクト

## 🎉 完成しました！

Tools28プロジェクトが、複数のRevitバージョンに対応した開発環境として完成しました。

---

## 📦 このパッケージに含まれるもの

### ソースコード
- `Commands/` - すべての機能コマンド
- `Resources/` - アイコンなどのリソース
- `Properties/` - アセンブリ情報
- `Application.cs` - メインアプリケーション
- `Tools28.csproj` - マルチバージョン対応プロジェクトファイル

### 自動化スクリプト
- `BuildAll.ps1` - すべてのバージョンを一括ビルド
- `GenerateAddins.ps1` - .addinファイルを生成
- `CreatePackages.ps1` - 配布パッケージ作成
- `Deploy-For-Testing.ps1` - テスト用デプロイ

### ドキュメント
- `README.md` - プロジェクト概要
- `QUICKSTART.md` - 5分でスタート
- `DEVELOPMENT_GUIDE.md` - 詳細な開発ガイド
- `ARCHITECTURE.md` - 技術的なアーキテクチャ

### 設定ファイル
- `.gitignore` - Git管理用の除外設定

---

## 🚀 今すぐ始める

### ステップ1: ZIPを解凍

ダウンロードした`Tools28_MultiVersion.zip`を任意の場所に解凍してください。

推奨場所：
- `C:\Projects\Tools28\`
- `C:\Users\[ユーザー名]\Documents\Projects\Tools28\`

### ステップ2: GitHubリポジトリを作成（推奨）

1. GitHubで新しいリポジトリを作成
   - リポジトリ名: `Tools28`
   - プライベート or パブリック（お好みで）

2. ローカルでGit初期化
```bash
cd Tools28
git init
git add .
git commit -m "Initial commit: Multi-version Revit addin"
git remote add origin https://github.com/[your-username]/Tools28.git
git branch -M main
git push -u origin main
```

### ステップ3: ビルドしてみる

```powershell
cd Tools28
.\BuildAll.ps1
```

これで5つのRevitバージョン（2022-2026）用のDLLが生成されます！

---

## 📁 生成されるファイル構造

ビルド後：
```
Tools28/
├── bin/
│   └── Release/
│       ├── Revit2022/
│       │   └── Tools28.dll
│       ├── Revit2023/
│       │   └── Tools28.dll
│       ├── Revit2024/
│       │   └── Tools28.dll
│       ├── Revit2025/
│       │   └── Tools28.dll
│       └── Revit2026/
│           └── Tools28.dll
│
├── Addins/                    # GenerateAddins.ps1 実行後
│   ├── Tools28_Revit2022.addin
│   ├── Tools28_Revit2023.addin
│   ├── Tools28_Revit2024.addin
│   ├── Tools28_Revit2025.addin
│   └── Tools28_Revit2026.addin
│
└── Packages/                  # CreatePackages.ps1 実行後
    ├── Tools28_Revit2022.zip
    ├── Tools28_Revit2023.zip
    ├── Tools28_Revit2024.zip
    ├── Tools28_Revit2025.zip
    └── Tools28_Revit2026.zip
```

---

## 🎯 主な機能

### 対応Revitバージョン
- ✅ Revit 2022
- ✅ Revit 2023
- ✅ Revit 2024
- ✅ Revit 2025
- ✅ Revit 2026

### 実装済みコマンド
1. **シート作成** - 複数シートの効率的な作成
2. **ビューコピー** - ビュー設定のコピー＆ペースト
3. **グリッドバブル** - グリッドバブルの制御
4. **セクションボックスコピー** - セクションボックス設定のコピー＆ペースト
5. **ビューポート位置調整** - ビューポート位置の簡単な調整
6. **クロップボックスコピー** - クロップボックス設定のコピー＆ペースト

---

## 🔄 典型的な開発フロー

### 職場PCでの作業
```bash
git pull                        # 最新を取得
# コード編集
.\BuildAll.ps1                  # ビルド
.\Deploy-For-Testing.ps1 -RevitVersion 2024  # テスト
git add .
git commit -m "機能追加: XXX"
git push
```

### 自宅PCでの作業
```bash
git pull                        # 最新を同期
# 開発継続...
```

---

## 💻 Claude Codeでの開発

### Claude Codeとは
ターミナルベースのAI開発アシスタント。コマンド一つで機能追加やバグ修正が可能。

### 使い方
```bash
cd Tools28
claude-code
```

プロンプト例：
```
「新しいコマンドを追加してください：
壁の高さを選択した壁すべてに一括設定する機能。
WPFダイアログで高さを入力できるようにしてください。」
```

Claude Codeが自動的に：
- ✅ ファイルとフォルダを作成
- ✅ コードを実装
- ✅ プロジェクトファイルを更新
- ✅ Application.csに登録
- ✅ ビルドして確認

---

## 📝 よくあるタスク

### 全バージョンをビルドする
```powershell
.\BuildAll.ps1
```

### 特定バージョンだけビルド
```powershell
msbuild Tools28.csproj /p:Configuration=Release /p:RevitVersion=2024
```

### テスト用にRevitへデプロイ
```powershell
.\Deploy-For-Testing.ps1 -RevitVersion 2024
```

### 配布パッケージを作成
```powershell
.\GenerateAddins.ps1
.\CreatePackages.ps1
```

### 新しい機能を追加
1. `Commands/`に新しいフォルダ作成
2. コマンドクラスを実装
3. `Tools28.csproj`にファイル追加
4. `Application.cs`にボタン登録
5. ビルド＆テスト

---

## 🆘 トラブルシューティング

### Q: ビルドが失敗する
```powershell
# PowerShellの実行ポリシーを変更
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Q: MSBuildが見つからない
Visual Studio Installerで「.NETデスクトップ開発」をインストール

### Q: Revitでアドインが読み込まれない
1. ファイルが正しい場所にあるか確認：
   `%APPDATA%\Autodesk\Revit\Addins\[version]\`
2. Revitを完全に再起動
3. Revitジャーナルファイルでエラー確認

### Q: 複数PCで開発したい
GitHubリポジトリを作成して、各PCで`git clone`

---

## 📚 次に読むべきドキュメント

1. **今すぐ始めたい** → `QUICKSTART.md`
2. **詳しい開発手順を知りたい** → `DEVELOPMENT_GUIDE.md`
3. **技術的な詳細を知りたい** → `ARCHITECTURE.md`

---

## 🎊 これで準備完了！

このプロジェクトは、以下のことが可能です：

✅ 単一のソースコードから5つのRevitバージョン対応DLLを生成
✅ 1コマンドで全バージョンをビルド
✅ 配布パッケージの自動作成
✅ GitHubで複数PC間での同期
✅ Claude Codeでの効率的な開発

Happy Coding! 🚀

---

## 📞 サポート

質問や問題がある場合：
- GitHubでIssueを作成
- ドキュメントを参照

開発を楽しんでください！
