# Tools28 - クイックスタートガイド

このガイドでは、Tools28の開発環境を最短で立ち上げる方法を説明します。

## 🚀 5分で始める

### ステップ1: プロジェクトを取得

```bash
# GitHubからクローン
git clone https://github.com/[your-username]/Tools28.git
cd Tools28
```

### ステップ2: すべてのバージョンをビルド

```powershell
# PowerShellで実行
.\BuildAll.ps1
```

完了！これで5つのRevitバージョン用のDLLが生成されます。

---

## 📦 配布パッケージを作成する

```powershell
# .addinファイルを生成
.\GenerateAddins.ps1

# ZIPパッケージを作成
.\CreatePackages.ps1
```

`Packages/`フォルダに配布用ZIPファイルが作成されます。

---

## 🧪 Revitでテストする

```powershell
# 例: Revit 2024でテストする場合
.\Deploy-For-Testing.ps1 -RevitVersion 2024
```

その後、Revit 2024を起動してテスト。

---

## 🔄 日常的な開発フロー

### 1. 最新コードを取得
```bash
git pull
```

### 2. コードを編集
お好みのエディタで編集

### 3. ビルド
```powershell
.\BuildAll.ps1
```

### 4. テスト
```powershell
.\Deploy-For-Testing.ps1 -RevitVersion 2024
```

### 5. コミット&プッシュ
```bash
git add .
git commit -m "機能追加: XXX"
git push
```

---

## 💻 Claude Codeを使う場合

```bash
# プロジェクトディレクトリで
claude-code
```

プロンプト例：
```
「新しいコマンド『WallHeightAdjuster』を追加してください。
選択した壁の高さを一括変更する機能です。」
```

Claude Codeが自動的に：
- ファイルを作成
- コードを実装
- プロジェクトを更新
- ビルドまで実行

---

## 🎯 よく使うコマンド一覧

| 目的 | コマンド |
|------|----------|
| すべてビルド | `.\BuildAll.ps1` |
| 特定バージョンビルド | `msbuild Tools28.csproj /p:RevitVersion=2024 /p:Configuration=Release` |
| .addin生成 | `.\GenerateAddins.ps1` |
| パッケージ作成 | `.\CreatePackages.ps1` |
| テストデプロイ | `.\Deploy-For-Testing.ps1 -RevitVersion 2024` |
| 最新取得 | `git pull` |
| 変更を保存 | `git add . && git commit -m "message" && git push` |

---

## 🆘 トラブルシューティング

### ビルドが失敗する
```powershell
# 実行ポリシーを変更
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Revitでアドインが表示されない
1. ファイルの場所を確認: `%APPDATA%\Autodesk\Revit\Addins\2024\`
2. Revitを完全に再起動
3. Revitのジャーナルファイルでエラーを確認

### Git pushが失敗する
```bash
# HTTPSで認証
git remote set-url origin https://github.com/[your-username]/Tools28.git
```

---

## 📚 さらに詳しく

詳細な情報は以下を参照：
- [README.md](README.md) - プロジェクト概要
- [DEVELOPMENT_GUIDE.md](DEVELOPMENT_GUIDE.md) - 詳細な開発手順

---

## 🎉 完了！

これでTools28の開発環境が整いました。
Happy Coding! 🚀
