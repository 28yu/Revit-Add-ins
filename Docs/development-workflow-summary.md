# Tools28 開発ワークフローまとめ

## 登場人物の相関図

### 登場人物（使用ソフト）の紹介

| 登場人物 | ドラマでの役割 | 説明 |
|---|---|---|
| **Claude Code** | 脚本家 | AIがコード（プログラムの設計図）を書いてくれるツール |
| **Visual Studio** | もう一人の脚本家 / 調査員 | コードを書いたり、デバッグ（不具合の原因を調査・修正すること）をするためのエディタ |
| **PowerShell** | 現場スタッフ | ビルド（コードを実行可能なファイルに変換すること）やデプロイ（完成したファイルを所定の場所に配置すること）の実作業を行う |
| **GitHub** | 台本倉庫 + 配達係 | コードを保管・共有するサービス。リリース（完成版を公開すること）も自動で行う |
| **NuGet** | 小道具係 | API（他のソフトの機能を借りるための窓口）の部品を自動で調達してくる |
| **Revit** | 舞台 | 開発したアドイン（追加機能）が実際に動く場所 |

### 相関図

```
                        【開発フェーズ】

    Claude Code                     Visual Studio
   （脚本家）                      （もう一人の脚本家）
        |                                |
        |  コードを書く                  |  コードを書く
        |                                |  デバッグ（不具合調査）する
        v                                v
   +---------------------------------------------------------+
   |                                                         |
   |              コード（プログラムの設計図）                |
   |                                                         |
   +---------------------------------------------------------+
        |                                          |
        |  保管する                                |
        v                                          |
     GitHub                                        |
   （台本倉庫）                                    |
                                                   |
                                                   v
                        【ビルド＆テストフェーズ】

                      PowerShell
                    （現場スタッフ）
                         |
          +--------------+--------------+
          |                             |
          v                             v
   ビルド（変換）する             デプロイ（配置）する
          |                             |
          |                             v
          |                          Revit
          |                        （舞台）
          |                       動作を確認する
          |
          |  部品を届ける
          |<-----------  NuGet
          |            （小道具係）
          |
          |
          v
                         【リリースフェーズ】

     GitHub  ──────>  GitHub Actions（自動スタッフ）
   （台本倉庫）            |
        ^                  |  全バージョンを自動で
        |                  |  ビルド（変換）＆ ZIP 作成
   タグ（バージョン        |
   番号の目印）を          v
   付けて送る         GitHub Releases
                     （完成品の受け渡し窓口）
                           |
                           v
                     ユーザーがダウンロード
```

### 関係のまとめ

- **Claude Code / Visual Studio** → コードを書く → **GitHub** に保管する
- **PowerShell** → コードをビルド（実行ファイルに変換）→ **Revit** にデプロイ（配置）する
- **NuGet** → ビルド時に Revit の API（機能の窓口）部品を **PowerShell** に届ける
- **GitHub** → タグ push をきっかけに自動でビルド → ZIP を作成 → **Releases** で公開する

> **初心者向けメモ**: NuGet はビルド時に裏で自動的に動くので、最初は意識しなくて OK です。

---

## 開発環境

- **言語**: C# / .NET Framework 4.8 (Revit 2021-2024) / .NET 8 (Revit 2025-2026)
- **対応バージョン**: Revit 2021, 2022, 2023, 2024, 2025, 2026
- **必要ツール**:
  - Visual Studio 2022 (.NET デスクトップ開発ワークロード)
  - .NET Framework 4.8 開発ツール / .NET 8.0 SDK
  - PowerShell 5.0+
  - Revit (テスト用、デフォルトは 2022)
- **Revit API**: NuGet パッケージ `Nice3point.Revit.Api` 経由で取得（ローカル Revit のインストール不要でビルド可能）

---

## 日常の開発サイクル

1. **コードを書く**
   - `Commands/機能名/` フォルダを作成し、`IExternalCommand` を実装
   - `Application.cs` の `OnStartup()` でリボンにボタンを登録
   - （任意）`Resources/Icons/` にアイコンを追加
2. **ビルド & デプロイ**
   - `.\QuickBuild.ps1` を実行（Revit 2022 用にビルド → Addins フォルダへ自動デプロイ）
   - 別バージョンで開発する場合は `dev-config.json` を編集、または `.\QuickBuild.ps1 -RevitVersion 2024` で指定
3. **Revit で動作確認**
   - Revit を起動（または再起動）し、「28 Tools」タブで新機能をテスト
4. **問題があれば修正して繰り返し**
   - デバッグ: Visual Studio から Revit.exe にアタッチ
   - ログ出力先: `C:\temp\Tools28_debug.txt`

---

## 新機能追加の手順

1. `Commands/` 配下に新しいフォルダとコマンドクラスを作成
   - `[Transaction(TransactionMode.Manual)]` 属性を付与
   - 名前空間: `Tools28.Commands.機能名`
2. `Application.cs` にリボンボタンを追加
3. （任意）32x32 PNG アイコンを `Resources/Icons/` に配置し、`.csproj` に `<Resource>` タグを追加
4. `.\QuickBuild.ps1` でビルド & テスト
   - SDK-style csproj のため `.cs` ファイルは自動認識（`<Compile Include>` 不要）

---

## リリース手順

1. **全バージョンをビルド**
   - `.\BuildAll.ps1` を実行 → 6バージョン分の DLL が `bin\Release\Revit{VERSION}\` に出力
2. **配布パッケージを作成**
   - `.\CreatePackages.ps1 -Version "X.X"` を実行
   - 出力: `Dist\28Tools_Revit{VERSION}_vX.X.zip`（各バージョン用 ZIP）
3. **動作確認**
   - 必要に応じて複数バージョンの Revit で検証
4. **コミット & プッシュ**
   - `git add .` → `git commit -m "..."` → `git push`
5. **GitHub Releases で公開**
   - `git tag vX.X` → `git push --tags`
   - GitHub Actions が自動で全バージョンのビルド → ZIP 作成 → Releases にアップロード

---

## 配布パッケージの構成

- ZIP ファイル名: `28Tools_Revit{VERSION}_vX.X.zip`
- 内容:
  - `28Tools/Tools28.dll` — メイン DLL
  - `28Tools/Tools28.addin` — マニフェストファイル
  - `install.bat` — ワンクリックインストール（管理者実行）
  - `uninstall.bat` — アンインストール
  - `README.txt` — インストール手順
- インストール先: `C:\ProgramData\Autodesk\Revit\Addins\{VERSION}\`

---

## CI/CD (GitHub Actions)

- **自動リリース**: タグ push (`git tag vX.X` → `git push --tags`) で起動
  - 全6バージョン (2021-2026) を自動ビルド
  - 配布 ZIP を作成し GitHub Releases にアップロード
- **手動実行**: GitHub Actions 画面 → "Build and Release" → Run workflow → バージョン番号を入力

---

## ビルドスクリプト一覧

| スクリプト | 用途 |
|---|---|
| `QuickBuild.ps1` | 開発用: 単一バージョンのビルド & Addins フォルダへデプロイ |
| `BuildAll.ps1` | リリース用: 全6バージョン一括ビルド |
| `CreatePackages.ps1` | 配布 ZIP 作成 |
| `GenerateAddins.ps1` | `.addin` マニフェストファイル生成 |
| `Deploy-For-Testing.ps1` | テスト用手動デプロイ |

---

## 注意事項

- トランザクションは `TransactionMode.Manual` を使用すること
- WPF ダイアログは XAML + コードビハインドで構成
- 条件付きコンパイルシンボル: `REVIT2021` ~ `REVIT2026` (バージョン別分岐に使用)
- プロジェクトファイル (`Tools28.csproj`) は SDK-style で、ターゲットフレームワークはバージョンに応じて自動切替え
