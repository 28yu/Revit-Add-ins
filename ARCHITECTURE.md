# Tools28 - アーキテクチャドキュメント

このドキュメントでは、Tools28の技術的なアーキテクチャと設計思想を説明します。

## プロジェクト構造

```
Tools28/
│
├── Commands/                          # コマンド実装（機能別フォルダ）
│   ├── CropBoxCopy/                  # クロップボックスコピー
│   │   └── CropBoxCopyCommands.cs
│   ├── GridBubble/                   # グリッドバブル制御
│   │   └── GridBubbleCommands.cs
│   ├── SectionBoxCopy/               # セクションボックスコピー
│   │   └── SectionBoxCopyCommands.cs
│   ├── SheetCreation/                # シート作成
│   │   ├── SheetCreationCommands.cs
│   │   ├── SheetCreationDialog.xaml
│   │   └── SheetCreationDialog.xaml.cs
│   ├── ViewCopy/                     # ビューコピー
│   │   └── ViewCopyCommands.cs
│   └── ViewportPosition/             # ビューポート位置調整
│       └── ViewportPositionCommands.cs
│
├── Resources/                         # リソースファイル
│   └── Icons/                        # アイコン画像（32x32 PNG）
│       ├── both_32.png
│       ├── left_32.png
│       ├── right_32.png
│       ├── sheet_creation_32.png
│       ├── view_copy_32.png
│       ├── view_paste_32.png
│       ├── sectionbox_copy_32.png
│       ├── sectionbox_paste_32.png
│       ├── viewport_copy_32.png
│       ├── viewport_paste_32.png
│       ├── cropbox_copy_32.png
│       └── cropbox_paste_32.png
│
├── Properties/                        # アセンブリ情報
│   └── AssemblyInfo.cs
│
├── Application.cs                     # メインアプリケーションクラス
├── Tools28.csproj                     # プロジェクトファイル（マルチバージョン対応）
│
├── BuildAll.ps1                       # ビルド自動化スクリプト
├── GenerateAddins.ps1                 # .addin生成スクリプト
├── CreatePackages.ps1                 # パッケージ作成スクリプト
├── Deploy-For-Testing.ps1             # テストデプロイスクリプト
│
├── .gitignore                         # Git除外設定
├── README.md                          # プロジェクト説明
├── QUICKSTART.md                      # クイックスタート
├── DEVELOPMENT_GUIDE.md               # 開発ガイド
└── ARCHITECTURE.md                    # このファイル
```

---

## マルチバージョン対応の仕組み

### 条件付きコンパイル

Tools28は、単一のソースコードから複数のRevitバージョンに対応したDLLを生成します。

#### プロジェクトファイル（Tools28.csproj）の主要設定

```xml
<!-- RevitVersionプロパティ（環境変数で指定） -->
<RevitVersion Condition="'$(RevitVersion)' == ''">2024</RevitVersion>

<!-- バージョン別の出力パス -->
<OutputPath>bin\Release\Revit$(RevitVersion)\</OutputPath>

<!-- バージョン別の定数定義 -->
<DefineConstants>TRACE;REVIT$(RevitVersion)</DefineConstants>

<!-- バージョン別のRevit API参照 -->
<ItemGroup Condition="'$(RevitVersion)' == '2024'">
  <Reference Include="RevitAPI">
    <HintPath>C:\Program Files\Autodesk\Revit 2024\RevitAPI.dll</HintPath>
  </Reference>
</ItemGroup>
```

#### ソースコードでのバージョン分岐

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

### ビルドプロセス

```
BuildAll.ps1 実行
    ↓
環境変数 RevitVersion を設定（2022, 2023, 2024...）
    ↓
MSBuild を各バージョンで実行
    ↓
出力: bin/Release/Revit2022/Tools28.dll
      bin/Release/Revit2023/Tools28.dll
      bin/Release/Revit2024/Tools28.dll
      ...
```

---

## コマンドアーキテクチャ

### 基本構造

各コマンドは以下の構造に従います：

```csharp
namespace Tools28.Commands.[CommandName]
{
    [Transaction(TransactionMode.Manual)]
    public class [CommandName]Command : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // 1. UIDocumentとDocumentを取得
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            // 2. トランザクション開始
            using (Transaction trans = new Transaction(doc, "[操作名]"))
            {
                trans.Start();
                
                // 3. Revit要素の操作
                // ...
                
                trans.Commit();
            }
            
            return Result.Succeeded;
        }
    }
}
```

### コマンドの種類

#### 1. 単純なコマンド（UIなし）
例: `CropBoxCopyCommands.cs`

```
ユーザー操作 → Revitデータ取得 → クリップボードに保存
                                    ↓
                          次回実行時にペースト
```

#### 2. ダイアログを持つコマンド
例: `SheetCreationCommands.cs`

```
ユーザー操作 → ダイアログ表示 → ユーザー入力 → Revit操作
```

---

## UIアーキテクチャ（WPF）

### ダイアログの構造

```
SheetCreationDialog.xaml          # UIデザイン（XAML）
        ↓
SheetCreationDialog.xaml.cs       # コードビハインド
        ↓
SheetCreationCommands.cs          # コマンド本体
```

### WPFダイアログの実装パターン

```csharp
// コマンド側
var dialog = new SheetCreationDialog();
bool? result = dialog.ShowDialog();

if (result == true)
{
    // ユーザーが入力した値を使用
    string prefix = dialog.SheetPrefix;
    // ... Revit操作
}
```

```csharp
// ダイアログ側（SheetCreationDialog.xaml.cs）
public partial class SheetCreationDialog : Window
{
    public string SheetPrefix { get; set; }
    
    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}
```

---

## Revitリボンの構成

### Application.cs の役割

`Application.cs`は、Revitアドインのエントリポイントです。

```csharp
public class Application : IExternalApplication
{
    public Result OnStartup(UIControlledApplication application)
    {
        // 1. リボンタブを作成
        application.CreateRibbonTab("Tools28");
        
        // 2. リボンパネルを作成
        RibbonPanel panel = application.CreateRibbonPanel(
            "Tools28", "Tools");
        
        // 3. ボタンを追加
        PushButtonData buttonData = new PushButtonData(
            "SheetCreation",
            "シート作成",
            Assembly.GetExecutingAssembly().Location,
            "Tools28.Commands.SheetCreation.SheetCreationCommand");
        
        // 4. アイコンを設定
        buttonData.LargeImage = GetEmbeddedImage("sheet_creation_32.png");
        
        panel.AddItem(buttonData);
        
        return Result.Succeeded;
    }
}
```

### リボンの構造

```
Revitリボン
└── Tools28タブ
    └── Toolsパネル
        ├── シート作成ボタン
        ├── ビューコピーボタン
        ├── ビューペーストボタン
        ├── セクションボックスコピーボタン
        ├── セクションボックスペーストボタン
        ├── ビューポート位置ボタン
        ├── クロップボックスコピーボタン
        ├── クロップボックスペーストボタン
        └── グリッドバブルボタン
```

---

## データフロー

### 例: ビューコピー機能

```
1. ユーザーがコピーボタンをクリック
        ↓
2. ViewCopyCommand.Execute() が呼ばれる
        ↓
3. アクティブビューの設定を取得
        ↓
4. 設定を静的変数に保存（クリップボード的な役割）
        ↓
5. TaskDialog で完了を通知
        ↓
（別のビューに移動）
        ↓
6. ユーザーがペーストボタンをクリック
        ↓
7. ViewPasteCommand.Execute() が呼ばれる
        ↓
8. 保存された設定を読み込み
        ↓
9. 現在のビューに設定を適用
        ↓
10. トランザクションをコミット
```

### クリップボードパターン

```csharp
public class ViewCopyCommands
{
    // 静的変数でデータを保持（セッション中有効）
    private static ViewSettings copiedSettings = null;
    
    [Transaction(TransactionMode.Manual)]
    public class CopyCommand : IExternalCommand
    {
        public Result Execute(...)
        {
            // コピー
            copiedSettings = GetCurrentViewSettings(view);
            return Result.Succeeded;
        }
    }
    
    [Transaction(TransactionMode.Manual)]
    public class PasteCommand : IExternalCommand
    {
        public Result Execute(...)
        {
            // ペースト
            if (copiedSettings != null)
            {
                ApplySettings(view, copiedSettings);
            }
            return Result.Succeeded;
        }
    }
}
```

---

## エラーハンドリング

### 基本的なパターン

```csharp
public Result Execute(...)
{
    try
    {
        // 処理
        return Result.Succeeded;
    }
    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
    {
        // ユーザーがキャンセルした場合
        return Result.Cancelled;
    }
    catch (Exception ex)
    {
        // その他のエラー
        message = ex.Message;
        return Result.Failed;
    }
}
```

### ユーザーへの通知

```csharp
// 成功メッセージ
TaskDialog.Show("成功", "操作が完了しました");

// エラーメッセージ
TaskDialog.Show("エラー", $"操作に失敗しました: {ex.Message}");

// 確認ダイアログ
TaskDialogResult result = TaskDialog.Show(
    "確認",
    "本当に実行しますか？",
    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
    
if (result == TaskDialogResult.Yes)
{
    // 処理を続行
}
```

---

## パフォーマンス最適化

### トランザクションの最適化

```csharp
// ❌ 悪い例: 複数のトランザクション
foreach (var element in elements)
{
    using (Transaction trans = new Transaction(doc, "Update"))
    {
        trans.Start();
        element.Update();
        trans.Commit();
    }
}

// ✅ 良い例: 単一のトランザクション
using (Transaction trans = new Transaction(doc, "Update All"))
{
    trans.Start();
    foreach (var element in elements)
    {
        element.Update();
    }
    trans.Commit();
}
```

### フィルタリングの最適化

```csharp
// ❌ 悪い例: すべての要素を取得してフィルタリング
var allElements = new FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
    .ToElements();
var walls = allElements.Where(e => e is Wall).ToList();

// ✅ 良い例: フィルターを使用
var walls = new FilteredElementCollector(doc)
    .OfClass(typeof(Wall))
    .WhereElementIsNotElementType()
    .ToElements();
```

---

## テスト戦略

### 単体テスト

現在、Tools28には自動テストは含まれていませんが、将来的には以下のような構造を推奨：

```
Tools28.Tests/
├── Commands/
│   ├── ViewCopyCommandsTests.cs
│   └── SheetCreationCommandsTests.cs
└── TestUtilities/
    └── MockRevitDocument.cs
```

### 手動テストチェックリスト

各コマンドに対して：

- [ ] 正常系: 期待通りに動作するか
- [ ] 異常系: 不正な入力でエラーハンドリングが機能するか
- [ ] 境界値: 極端な値（0, 最大値など）で動作するか
- [ ] UI: ダイアログが正しく表示されるか
- [ ] パフォーマンス: 大量のデータで問題ないか

---

## セキュリティとベストプラクティス

### トランザクション管理

- 常に`using`ステートメントでトランザクションを管理
- 例外が発生した場合は自動的にロールバック
- `trans.Commit()`を明示的に呼ぶ

### リソース管理

- WPFダイアログは`ShowDialog()`で適切に破棄される
- 画像リソースは埋め込みリソースとして管理

### エラーメッセージ

- ユーザーフレンドリーなメッセージ
- 技術的な詳細は開発者向けログに記録
- スタックトレースはエンドユーザーに見せない

---

## 拡張性

### 新しいコマンドの追加手順

1. `Commands/`に新しいフォルダを作成
2. コマンドクラスを実装
3. 必要に応じてダイアログ（XAML）を作成
4. `Tools28.csproj`にファイルを追加
5. `Application.cs`にリボンボタンを登録
6. アイコンを`Resources/Icons/`に追加

### プラグインアーキテクチャへの拡張

将来的には以下のような拡張も可能：

```csharp
// プラグインインターフェース
public interface ITools28Plugin
{
    string Name { get; }
    void Initialize(UIControlledApplication app);
}

// プラグイン読み込み
private void LoadPlugins()
{
    var pluginDir = Path.Combine(AssemblyDirectory, "Plugins");
    // プラグインDLLをロードして初期化
}
```

---

## 参考資料

- [Revit API ドキュメント](https://www.revitapidocs.com/)
- [Revit API Developer's Guide](https://help.autodesk.com/view/RVT/2024/ENU/?guid=Revit_API_Revit_API_Developers_Guide_html)
- [WPF ドキュメント](https://docs.microsoft.com/dotnet/desktop/wpf/)

---

このドキュメントは、Tools28の技術的な設計と実装の詳細を記録しています。
新しい機能を追加する際や、アーキテクチャを変更する際は、このドキュメントも更新してください。
