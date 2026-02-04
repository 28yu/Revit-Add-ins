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
├── Packages/                   # 配布パッケージ (バージョン別)
│   ├── 2021/ ~ 2026/           # 各バージョン用 install.bat, uninstall.bat, README.txt
├── BuildAll.ps1                # 全バージョン一括ビルド
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

### 構成 (Packages/{VERSION}/)
各バージョン (2021-2026) に以下のファイルを格納:
- `install.bat` - 管理者実行でアドインをインストール
- `uninstall.bat` - アドインをアンインストール
- `README.txt` - ユーザー向けインストール手順

### インストール先
`C:\ProgramData\Autodesk\Revit\Addins\{VERSION}\`

### 配布ZIP作成手順
```powershell
.\GenerateAddins.ps1
.\CreatePackages.ps1
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

## 注意事項

- Revit API の `RevitAPI.dll` / `RevitAPIUI.dll` は `Private=False` で参照 (ローカルコピーしない)
- トランザクションは `TransactionMode.Manual` を使用
- デバッグログは `C:\temp\Tools28_debug.txt` に出力
- WPFダイアログを使用するコマンドは XAML + コードビハインドで構成
