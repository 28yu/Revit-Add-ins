# 型枠数量算出アドイン 開発状況ハンドオフ

**最終更新**: 2026-04-27
**開発用 Revit**: 2022 (`dev-config.json`)

---

## 現在の状態（2026-04-27 セッション終了時点）

### 動作確認済み・本番投入可能な機能
- 6カテゴリ（構造柱・構造フレーム・壁・床・構造基礎・階段）の型枠面積算出
- 接触面の自動控除（Full Contact + Partial Contact 両対応）
- DirectShape による色分け視覚化
- 3D解析ビュー「型枠分析」自動生成
- 集計表「型枠数量集計」自動生成（レベル → 部位 → タイプ名で階層グループ化）
- 面積フィールドに合計を計算表示
- Excel エクスポート

### このセッションで対応した改善項目
| # | 内容 | 関連ファイル |
|---|---|---|
| 1 | 集計表の面積フィールドに「合計を計算」を設定 | `Output/ScheduleCreator.cs` |
| 2 | 3D解析ビューでセクションボックス枠線・レベル線を非表示 | `Output/FormworkVisualizer.cs` |
| 3 | 3D解析ビューの視点を実行時のアクティブ3Dビューに合わせる | `Output/FormworkVisualizer.cs`, `FormworkCalculatorCommand.cs` |

### 開発時に判明した重要な API 知見

#### Revit 2022 では `ScheduleField.HasTotals` が存在しない
- 過去のセッションでリフレクション経由で `HasTotals = true` を試みていたが、Revit 2022 の API では公開プロパティとして存在しないため無効化していた（プロパティ取得結果が null）。
- **正解 API**: `ScheduleField.DisplayType = ScheduleFieldDisplayType.Totals`
- `ScheduleFieldDisplayType` enum 値（Revit 2022 で確認済み）:

| 値 | UI ラベル |
|---|---|
| `Standard` (0) | 計算しない |
| **`Totals` (1)** | **合計を計算** |
| `MinMax` (2) | 最小値と最大値を計算 |
| `Max` (3) | 最大値を計算 |
| `Min` (4) | 最小値を計算 |

`DisplayType` は Revit 2021-2026 全バージョンで公開 API として存在するため、リフレクション不要で直接プロパティアクセス可能。

#### View3D の視点コピー
- 既存ビューの視点を新規ビューにコピーするには `view.SetOrientation(sourceView.GetOrientation())` を使う
- アクティブビューが `View3D` でないとき（平面ビュー等）は何もしない（既定のアイソメトリックを維持）

#### View3D で整理すべきカテゴリ
- セクションボックス枠線: `BuiltInCategory.OST_SectionBox`
- レベル線: `BuiltInCategory.OST_Levels`
- どちらも `view.SetCategoryHidden(catId, true)` で非表示化（実行前に `CanCategoryBeHidden` でガード）

---

## 🎯 今後の機能拡張: マテリアルベース算出

### 背景と課題
現在は「カテゴリ」（構造柱・梁・壁・床等）を判定軸にして型枠を拾い出している。しかし実プロジェクトでは:
- カテゴリは同じでも材料が違う（例: RC柱 vs 鉄骨柱、RC壁 vs ALC壁）
- ALC や乾式間仕切等は型枠が不要
- カテゴリで一括処理すると、型枠不要な要素まで対象になってしまう
- 別途フィルタする手間がかかる

### 拡張方針: マテリアルから型枠数量を算出
要素のマテリアル（材料）を判定軸に変更／併用できるようにする。

#### 想定する要件
1. **マテリアル選択 UI**: ダイアログでプロジェクト内の全マテリアル一覧から複数選択
   - 例: `コンクリート - 現場打ち`, `コンクリート - プレキャスト` をチェック → これらの材料を持つ要素のみ対象
   - 「カテゴリで選ぶ」「マテリアルで選ぶ」「両方使う（AND/OR）」を切替可能

2. **判定対象**:
   - **構造材** (`Structural Material` パラメータ) を第一候補とする
   - 複合構造（壁・床）の場合は層ごとにマテリアルが違うので、**主構造層（Core）のマテリアル**を取得
   - フォールバック: タイプの「マテリアル」パラメータ → なければインスタンスのマテリアル

3. **算出ロジック**:
   - 既存の面分類・接触検出ロジックは変更不要（要素フィルタ部分のみ拡張）
   - `ElementCollector` に `IncludedMaterials` (`List<ElementId>`) フィールドを追加
   - `IncludedCategories` と AND/OR で組み合わせ可能に

4. **集計・色分け**:
   - 集計表の階層グループに「マテリアル」を追加できるオプション
   - 色分け基準にも「マテリアル別」を追加（既存の Category/Zone/FormworkType に並べる）

### 実装の見積もり

| 作業 | 影響範囲 | 工数感 |
|---|---|---|
| マテリアル選択 UI（ダイアログにリストボックス追加） | `Views/FormworkDialog.xaml(.cs)` | 中 |
| `FormworkSettings` にマテリアル関連フィールド追加 | `Models/FormworkSettings.cs` | 小 |
| 要素のマテリアル取得関数 | `Engine/ElementCollector.cs`（新規ヘルパー） | 中 |
| 要素フィルタにマテリアル条件を追加 | `Engine/ElementCollector.cs` | 中 |
| 共有パラメータにマテリアル名追加 | `Engine/FormworkParameterManager.cs` | 小 |
| 集計表にマテリアル列を追加（オプション） | `Output/ScheduleCreator.cs` | 小 |
| 色分けにマテリアル別を追加 | `Output/FormworkVisualizer.cs`, `Engine/FormworkFilterManager.cs` | 中 |
| 多言語化エントリ追加 | `Localization/Strings*.cs` (3ファイル) | 小 |
| `Models/FormworkSettings.cs` の `ColorSchemeType` enum 拡張 | `Models/FormworkSettings.cs` | 小 |

合計: 中規模（1〜2日相当）

### 設計時の注意点

#### マテリアル取得の優先順位
要素のマテリアル取得は意外と複雑なので、以下の順序でフォールバックさせる：
1. `Element.StructuralMaterialId`（構造材のみ。柱・梁・基礎で有効）
2. 複合構造の場合: `WallType.GetCompoundStructure().GetMaterialId(coreLayerIndex)`（壁・床）
3. タイプパラメータ `Material` (BuiltInParameter `MATERIAL_ID_PARAM`)
4. インスタンスパラメータでマテリアルが指定されている場合
5. ジオメトリの `Solid` から `Face.MaterialElementId` を取る（最終手段）

#### マテリアル名は表示名 (`Material.Name`) を使う
- ElementId は別ファイルにすると一致しない
- 名前ベースで集計すると、同名の異なる ElementId のマテリアルがある場合に統合されてしまう点に注意（その場合は ElementId 単位で集計し名前は表示用とする）

#### 後方互換性
- 既存の `IncludedCategories` 設定は維持
- マテリアル機能は新規オプトインフィールドとして追加（既存ユーザーの設定を壊さない）

---

## 既知の課題（未解決のまま終了）

### [18]-3 接触面検出の漏れ（一部ケース）
壁の T 字結合や梁と柱の取り合い部の一部ケースで、接触面の検出が漏れる場合がある。
- 第7世代「幾何学的検査 + UV内部判定 + 面積比 + Partial Contact + Boolean Difference」で大半は解決
- 残るのは特殊形状（Join Geometry で複雑に結合された要素、曲面を持つ面 など）
- リリース前に再評価が必要なら `C:\temp\Formwork_debug.txt` のログから攻める

### Phase 1/2 で実装済みのフォールバック
非対応ケースは Phase 1 の半透明 DirectShape で動作するためリグレッションはしない:
- 開口付き壁（CurveLoop が穴を含む）
- カーブウォール
- 非矩形の面
- UV投影失敗

---

## デバッグログ
- 出力先: `C:\temp\Formwork_debug.txt`
- 制御フラグ: `FormworkSettings.EnableDebugLog`（デフォルト `true`、UI 非露出）
- 上限 200,000 行（超えたら `... truncated` で停止）
- **リリース時は `false` に変更すること**

---

## 現在のコード構造

```
Commands/FormworkCalculator/
├── FormworkCalculatorCommand.cs       # エントリポイント
├── HANDOFF.md                         # このファイル
├── Models/
│   ├── FormworkSettings.cs            # UI設定（Scope/Categories/Grouping/Color等）
│   └── FormworkResult.cs              # 計算結果・面情報
├── Engine/
│   ├── ElementCollector.cs            # 要素収集 + カテゴリ判定
│   ├── SolidUnionProcessor.cs         # Solid 取得 + Boolean Union
│   ├── FaceClassifier.cs              # 面分類 (Top/Bottom/Required等)
│   ├── ContactFaceDetector.cs         # 接触面検出（Full + Partial）
│   ├── PartialContactClipper.cs       # 矩形ベース 2D 差分（Phase 2）
│   ├── SpatialGrid.cs                 # 空間索引（O(N²)→O(N) 相当）
│   ├── OpeningProcessor.cs            # 開口部処理
│   ├── FormworkCalcEngine.cs          # メインエンジン (3-Pass)
│   ├── FormworkParameterManager.cs    # 共有パラメータ管理
│   ├── FormworkFilterManager.cs       # View Filter 管理
│   └── FormworkDebugLog.cs            # デバッグログ
├── Output/
│   ├── ExcelExporter.cs               # Excel 出力
│   ├── FormworkVisualizer.cs          # 3Dビュー + DirectShape + 視点コピー
│   └── ScheduleCreator.cs             # 集計表作成（DisplayType=Totals）
└── Views/
    ├── FormworkDialog.xaml            # メインダイアログ
    └── FormworkDialog.xaml.cs
```

### 3 Pass パイプライン (`FormworkCalcEngine.Run`)
```
Pass 1: 要素毎に Solid 取得 → FaceClassifier で分類
Pass 2: ContactFaceDetector で接触面を DeductedContact に変更
Pass 3: 開口加算 + ElementResult 作成 + Aggregate
```

### FaceType enum
- `FormworkRequired`: 型枠必要
- `DeductedTop`: 最上面（スラブは全上向き面）
- `DeductedBottom`: 最下面（基礎のみ、それ以外は FormworkRequired にコンバート）
- `DeductedContact`: 他要素との接触面
- `DeductedBelowGL`: GL 以下
- `Inclined`: 傾斜面（現状未使用、全て FormworkRequired 扱い）
- `Error`: エラー

### 共有パラメータ (OST_GenericModel にバインド)
- `28Tools_FormworkMarker` (Text): DirectShape 識別マーカー = `"28Tools_Formwork"`
- `28Tools_Formwork_部位` (Text): 柱/梁/壁/スラブ/基礎/階段
- `28Tools_Formwork_レベル` (Text): 参照レベル名
- `28Tools_Formwork_区分` (Text): 色分けグループキー
- `28Tools_Formwork_面積` (Area): 面積 (㎡)
- `28Tools_Formwork_部分接触` (Text "Yes"/"No"): 一部消されている面の識別

### ビュー・集計表
- 解析 3D ビュー名: `型枠分析`（再実行で上書き、視点はソース3Dビューを継承）
- 集計表名: `型枠数量集計`（再実行で上書き）
- Excel 初期名: `型枠数量集計.xlsx`
- 集計表のグループ化: レベル → 部位 → タイプ名（ShowHeader/ShowFooter）
- `IsItemized=true`（インスタンス内訳）+ 面積フィールド `DisplayType=Totals`（合計表示）

### 色分け
- `FormworkFilterManager` で View Filter ベース
- フィルタルール: `28Tools_Formwork_区分 == <groupKey>`
- 元躯体: RGB(94,94,94) + 20% 透過のオーバーライド
- 解析ビュー以外では DirectShape を `View.HideElements()` で一括非表示

---

## ビルド・デプロイ
- 開発時: `QuickBuild.ps1`（Revit 2022 のみ）
- 全バージョン: `BuildAll.ps1`
- 自動デプロイ: ローカル `AutoBuild.ps1` が main の更新を検知して自動再ビルド＆デプロイ
- 詳細は CLAUDE.md 参照
