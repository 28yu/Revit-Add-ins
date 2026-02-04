# Tools28 - Revit Add-in 開発ガイド (Claude Code用)

## プロジェクト概要
- **名前**: Tools28
- **種類**: Autodesk Revit アドイン (C# / .NET Framework 4.8)
- **対応バージョン**: Revit 2021, 2022, 2023, 2024, 2025, 2026
- **名前空間**: `Tools28`
- **アセンブリ名**: `Tools28`
- **プロジェクトファイル**: `Tools28.csproj` (MSBuild, マルチバージョン対応)

## リポジトリ構成

```
Revit-Add-ins/
├── Application.cs              # メインアプリ (IExternalApplication) - リボンUI構築
├── Tools28.csproj              # マルチバージョン対応プロジェクトファイル
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
├── BuildAll.ps1                # 全バージョン一括ビルド (2021-2026)
├── GenerateAddins.ps1          # .addinマニフェスト生成
├── CreatePackages.ps1          # 配布ZIP作成
└── Deploy-For-Testing.ps1      # テスト用デプロイ
```

## ビルド

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

### テンプレート構成 (Packages/{VERSION}/)
各バージョン (2021-2026) に以下を格納:
```
Packages/{VERSION}/
├── 28Tools/
│   └── Tools28.addin    # マニフェスト (全バージョン共通内容)
├── install.bat           # Addins\{VERSION} へ DLL/addin をコピー
├── uninstall.bat         # Addins\{VERSION} から DLL/addin を削除
└── README.txt            # ユーザー向けインストール手順
```

### 配布ZIP構成 (CreatePackages.ps1 が生成)
```
28Tools_Revit{VERSION}_vX.X.zip
├── 28Tools/
│   ├── Tools28.dll       # ビルド成果物 (自動コピー)
│   └── Tools28.addin     # マニフェスト
├── install.bat
├── uninstall.bat
└── README.txt
```

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

## 新機能追加手順

1. `Commands/` に新しいフォルダを作成
2. `IExternalCommand` を実装するコマンドクラスを作成
3. `Tools28.csproj` の `<Compile Include="...">` にファイルを追加
4. `Application.cs` のリボンにボタンを登録
5. アイコンが必要な場合は `Resources/Icons/` に32x32 PNGを追加し `.csproj` に `<Resource>` を追加

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
