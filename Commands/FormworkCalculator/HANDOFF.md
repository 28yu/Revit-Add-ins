# 型枠数量算出アドイン 開発状況ハンドオフ

**最終更新**: 2026-05-07
**開発用 Revit**: 2022 (`dev-config.json`)

---

## 🎯 次セッションへの引き継ぎ事項（2026-05-07 終了時点）

### 動作確認済み・本番投入可能な機能
1. 6 カテゴリ（柱・梁・壁・床・基礎・階段）の型枠面積算出
2. 鉄骨部材の自動除外（4 層フォールバック判定）
3. デッキスラブの自動除外（タイプ名 "DS" 検出）
4. 壁スイープ・リビールの自動除外（contact detection 後に Excluded へ移動 + WallSweepFaceDeductor で host 壁面を直接控除）
5. 集計表「型枠数量集計」 + 動的合計サマリ「型枠数量集計_合計」
6. Excel 出力（CJK 対応）+ 解析 3D ビュー
7. 鉄骨除外 / デッキスラブ除外 / 壁スイープの可視化（オレンジ色 DirectShape、フィルタ既定 OFF）
8. パラメータ自動候補 ComboBox（工区・型枠種別）

### ⚠️ 未解決の課題（次セッションで取り組む）

#### [53] ContactFaceDetector の精度問題
ユーザー報告の症状（ログから具体例 E3280907, E3280909, E3707261 等を確認済み）:

1. **柱の一部接触で全面除外** (E3280907 等) — Stage 1 が過大マッチ
2. **構造フレームの T 字接合で接触部が省かれない** — Stage 2 検出漏れ
3. **構造基礎の床接触面が残る** — contact detection 漏れ
4. **構造基礎の梁接触が半透明だが面積が引かれない** — Partial 検出は OK だが Clipper 失敗
5. **部分接触 Yes で面積 0** — Clipper 失敗時の naive sum 過大評価（5% キャップで暫定対処済みだが根本解決ではない）

**最重要ケース**: 柱 E3280907 (950×950×940mm = 高さ 940mm のスタブ柱)
- 4 側面 (各 0.89 m²) が梁 E3286646 の側面 (各 0.77 m²) と Stage 1 (Full Contact) マッチ
- 面積比 1.16 < 1.2 (`AreaRatioLimit` 厳格化後でもまだ通る)
- 距離 d = 0 (中心投影が距離 0)
- ペアログ抜粋:
```
[Pair E3280907(Column) x E3286646(Beam)]
  f1->f1 dot=-1.000 aArea=9.6122 bArea=8.2828 d=0.0000 uv=(1.542,1.558) FULL_CONTACT
  f2->f5 ... d=0.0000 ... FULL_CONTACT
  f3->f6 ... d=0.0000 ... FULL_CONTACT
  f4->f7 ... d=0.0000 ... FULL_CONTACT
```
スクリーンショットでは梁の上に乗る短いスタブ柱で、4 側面は本来露出 (formwork 必要) のはず。
誤判定の仮説: **梁が L 字 / U 字でスタブ柱を 4 方向から囲んでいる** か、**Revit Join Geometry の影響で柱側面と梁側面が幾何的に「距離 0」になっている**。

**修正方針候補（次セッションで検討・実装）**:
- Stage 1 の判定強化: 距離 0 + 面積比だけでなく、a / b の UV 投影**重なり領域**もチェック
  - 例: a の 4 隅を b に投影して、b 内に何点入るか。3-4 点入った場合のみ Full Contact
  - 既に `ProjectBCornersToA` という関数で b → a の投影は行っているので、対称的に a → b も実装
- それでも誤検出する場合: 体積 (Solid Volume) 比較で「片方が他方に内包されているか」を確認
- スタブ柱と梁の組合わせは特殊ケースなので、専用ロジックの可能性も

#### [52-1] 壁の天端リビール切り取り箇所
- `WallSweepFaceDeductor` で host 壁面を BB 内中心点判定で `DeductedContact` 化済み
- ユーザーは「まだ改善されていない」と報告 (要再確認、もしかすると修正後のビルド未反映状態だった可能性)

### 直近の修正コミット (順)
| 日時 | 内容 |
|---|---|
| 2026-05-07 | 鉄骨除外 (4層判定) |
| 2026-05-07 | デッキスラブ除外 + 除外フィルタ既定非表示 |
| 2026-05-07 | Excel CJK 対応列幅 |
| 2026-05-07 | 集計表総合計を Excel と一致 (DirectShape area 要素単位集約) |
| 2026-05-07 | 動的合計サマリ集計表 + 集計後の自動アクティブ化 |
| 2026-05-07 | doc.Regenerate() で 3D ビュー描画問題解決 |
| 2026-05-07 | 壁スイープ除外 (contact detection 参加 → 後段で除外) |
| 2026-05-07 | 部分接触の Clipper ベース面積計算 |
| 2026-05-07 | WallSweepFaceDeductor + Stage 1 area ratio 1.5→1.2 |
| 2026-05-07 | ElemDiag / FaceDiag 詳細診断ログ追加 |
| 2026-05-07 | ElemDiag マーカー精緻化 + 寸法情報追加 |

### 診断ログの読み方
`C:\temp\Formwork_debug.txt` 内：

- **`[ElemDiag]`** — 1 要素 1 行のサマリ
  - `dim=950x950x940mm formwork=0.00m² faces=0/1/0/5/0/0 parts=0 ⚠️ZERO_EMBEDDED`
  - faces = `Required/Top/Bottom/Contact/BelowGL/Inclined`
  - ⚠️ マーカー: `ZERO`, `ZERO_EMBEDDED`, `CONTACT_HEAVY`, `ALL_PARTIAL`

- **`[FaceDiag]`** — 部分接触のある面ごとの詳細
  - `[FaceDiag] E12345 face[3] area=2.5000 partials=2 [ E22:a=0.4 E23:a=0.7 ] clipper-OK eff=1.4000`
  - Clipper 成否、フォールバック時の rawSum、最終 effective area

- **`[Pair E<a> x E<b>]`** — 要素ペアごとの面評価
  - `f1->f1 dot=-1.000 aArea=9.61 bArea=8.28 d=0.0 uv=(1.5,1.5) FULL_CONTACT`
  - `REJECTED(s1=理由,s2=理由)` で除外条件を確認

ユーザーが疑わしい要素 ID を `grep "⚠️"` で抽出 → `grep "Pair E<id>"` でペア評価を確認可能。

---

## 2026-05-07 セッション最終 (本日のまとめ): 重要な API 知見と最終仕様

### ⚠️ Revit 2022 Schedule API の重要な制約（再現したい場合の参考）

#### 1. `TableSectionData.SetCellStyle` は限定的
スタイルを上書きできるのは:
- ✓ Header セクションの全セル
- ✓ Body セクションの **行 0（列ヘッダー）のみ**
- ✗ Body セクションのデータ行・グループフッタ・**総合計行はスタイル変更不可**

エラーメッセージ:
```
ArgumentException: Only allow to override cell style for header section or
column header in body section.
```

#### 2. `TableCellStyle.FontSize` は Revit 2022 に存在しない
- Revit 2024+ で追加された
- リフレクションで設定しようとしても silently fail
- Revit 2022 でフォントサイズをプログラム経由で変更することは API 制約上**不可能**

#### 3. Schedule View にテキストスタイル参照パラメータが存在しない (Revit 2022)
- "Body text" / "Header text" / "Title text" 相当のパラメータは
  Schedule View にも、その Type 要素にも存在しない
- 確認方法: 全 ElementId 型パラメータを列挙しても "新しいビューに適用される
  ビューテンプレート" のみ
- → スケジュールのフォントサイズ変更は手動 (プロジェクトのテキストタイプ
  「文字サイズ」を変更) でのみ可能

#### 4. `ScheduleDefinition.GrandTotalTitle` は **存在するが要前提条件**
- Revit 2022 でも property は存在
- ただし `ShowGrandTotalTitle = true` を**先に設定しないと setter が
  TargetInvocationException を投げる**
- リフレクションで InnerException を取得して原因特定可能

#### 5. ⚠️ `doc.Regenerate()` を呼ばないと新規 DirectShape のジオメトリは認識されない
- 大きな落とし穴。`DirectShape.SetShape` 直後に `get_BoundingBox(null)` を
  呼んでも `null` が返る
- `Element.get_Geometry(opts)` も Solid 数 0 を返す
- 結果として **3D ビューに描画されない**
- 対策: 全 DirectShape 作成後に **`doc.Regenerate()`** を必ず呼ぶ
- 同じトランザクション内で OK

#### 6. `TableSectionData.SetColumnWidth` の単位は **feet (内部単位)**
- `0.167 ft ≈ 50mm`、`0.5 ft ≈ 152mm`
- タイトル「<…>」の改行は body 全体の幅で決まるため、列幅を広げると
  タイトルも改行しなくなる

#### 7. `ScheduleSheetInstance` ではなく `ViewSchedule.GetTableData()` 経由
- セクションデータは `tableData.GetSectionData(SectionType.Header/Body/Footer)`

---

### 動的合計サマリ集計表のパターン
ユーザー操作 (DirectShape 削除) に追従する合計値を styled に表示するレシピ:

1. メイン集計表とは別に「型枠数量集計_合計」を新設
2. `IsItemized = false` で全件を 1 行に集約
3. 件数 + 面積（`DisplayType=Totals`）の 2 フィールド
4. マーカーフィルタ (Equal MarkerValue) で formwork のみ
5. **Body 行 0（列ヘッダー）の各セルにスタイル設定** (赤字・太字・薄黄背景)
   ← 動的に追従する値の直上にラベルを置く構図
6. データ行 (Body 行 1) は Revit が DirectShape の追加・削除に応じて自動再計算
7. 列幅は 0.167 ft 程度に絞ってタイトル改行を防ぐ

→ ラベルは静的だが、その直下の値はリアルタイム追従する

---

### Excel 出力の CJK 対応
`ClosedXML.AdjustToContents()` は半角文字を 1 として幅計算するため、
日本語が見切れる。対策:
- `MeasureWidth(string)` ヘルパー: CJK 全角を 2.0、半角を 1.1 でカウント
- 各列の最大幅を直接 `Column.Width` にセット
- オートフィルタ付きシート (要素明細) は padding +5～8 文字必要
  （ドロップダウン矢印 ≈ 17px = 約 2.5 文字幅 + 余白）

`cell.Value` は `XLCellValue` (構造体) なので `?.` 演算子使用不可。
`cell.Value.ToString()` を使う。

---

### パラメータ自動候補 ComboBox（工区別・型枠種別）
`Engine/ParameterCandidateScanner.cs`:
- ParameterBindings (プロジェクト/共有パラメータ) + 主要カテゴリの先頭
  3 件のインスタンス・タイプから収集
- キーワード:
  - 工区: `工区 / ゾーン / Zone / エリア / Area / ブロック / Block / 区分 / Section / 範囲 / Phase`
  - 型枠種別: `型枠 / 種別 / Formwork / Type / パターン / Pattern / 仕様 / Spec`
- ComboBox は `IsEditable="True"` で手入力もサポート

---

### 鉄骨除外の 4 層フォールバック判定 (確定仕様)
`Engine/SteelMemberDetector.cs`:

| Layer | 判定 | 例 |
|---|---|---|
| L1 | `FamilyInstance.StructuralMaterialType == Steel` | 標準ファミリ |
| L2 | 断面形状分析（中空 or 充実率<0.5） | CFT、H形鋼 |
| L3 | 構造材マテリアル名 (Steel/鋼/鉄/Metal) | マテリアルのみ正設定 |
| L4 | ファミリ・タイプ名キーワード | 古いファミリ、CFT- 等 |

- 対象: 構造柱・構造フレームのみ
- SRC柱は中実・凸で保持、CFT柱は中空または "CFT" 名で除外
- 検出失敗時は **保持側にフェイルセーフ** (誤除外回避)
- 暗黙挙動 (UI 非露出)

### デッキスラブ除外
`Engine/DeckSlabDetector.cs`:
- 床のタイプ名 (`ElementType.Name`) または要素名に "DS" / "ＤＳ" を含めば除外
- `Contains` ベース（"DS150"・"ALC-DS" 等を拾う）
- 大文字のみ ("ds" は除外しない）

### 壁スイープ・リビール除外
`Engine/ElementCollector.cs` (CollectAndClassify):
- `WallSweep` クラスのインスタンス（壁スイープ・リビール両方を含む）を一律除外
- 壁の天端付帯部 (コーピング・水切り等) は別工法で施工されることが多く型枠不要
- ラベル: `WallSweepExcludedLabel = "壁スイープ(除外)"`

---

### 視覚化の最終仕様
- formwork DirectShape 厚み: **0.05 ft（≈15mm）**
- 元躯体: **50% 透過 + RGB(94,94,94) グレー**
- View Filter で色分け（区分パラメータベース）
- 除外フィルタは既定で非表示（V/G で手動 ON で確認）
- `OST_GenericModel` カテゴリは明示的に表示状態に設定
- `View3D.DisplayStyle = Shading` を明示
- formwork DirectShape は `SurfaceTransparency = 0` で完全不透明

### DirectShape の面積パラメータ集約 (Excel 総括表との一致)
- 要素単位の最終 `er.FormworkArea` (開口控除・端面加算反映) を最初の
  **FormworkRequired** DirectShape 一つにまとめて持たせる
- 残りの DirectShape は面積 0 m²
- 集計表の総合計 = Excel 総括表の合計 と完全一致

### 完了時のビュータブ展開
1. 3D ビュー (タブを開く)
2. サマリ集計表 (最終アクティブ → フォアグラウンド)
- メイン集計表は Project Browser から手動で開く

---

## 2026-05-07 セッション後半: デッキスラブ除外と除外フィルタ既定非表示

- 床カテゴリのうち**タイプ名に "DS" を含む**ものをデッキスラブとして自動除外
  - `Engine/DeckSlabDetector.cs`（`Floor.GetTypeId() → ElementType.Name` を検査、半角 "DS" / 全角 "ＤＳ"）
- 解析3Dビューの**除外フィルタは既定で非表示**（チェック OFF）
  - `FormworkFilterManager.ApplyColorFilters` で `key == ExcludedGroupKey` のときのみ `SetFilterVisibility(false)`
  - ユーザーが手動で ON にすると除外要素のオレンジ表示を確認できる
- 除外概念を一般化（鉄骨専用 → 鉄骨＋デッキスラブ）
  - `ExcludedSteelResult` → `ExcludedResult` に名称変更（`Kind` enum を追加: `Steel` / `DeckSlab`）
  - `MarkerValueSteel` → `MarkerValueExcluded` (`28Tools_Formwork_Excluded`)
  - `SteelExcludedGroupKey` → `ExcludedGroupKey` (値: `"除外"`、フィルタ名 `"型枠_除外"`)
  - 部位ラベルは Kind 別: `SteelExcludedLabel = "鉄骨(除外)"` / `DeckSlabExcludedLabel = "デッキスラブ(除外)"`
- `CleanupExistingFormworkShapes` を `StartsWith("28Tools_Formwork")` ベースに変更
  （旧 `28Tools_Formwork_Steel` マーカーも自動回収）
- 完了ダイアログを多項目化（鉄骨件数 / デッキスラブ件数 / フィルタ説明）

### 検出パターン (デッキスラブ)
- `Floor.GetTypeId()` → `ElementType.Name` をチェック
- ヒット条件: 半角 "DS" を **`Contains`** で検出（"DS150"・"Deck-DS"・"ALC-DS" 等を全て拾う）
- 全角 "ＤＳ" もカバー
- 大文字限定（"ds" は対象外）

---

## 2026-05-07 セッション: 鉄骨部材の自動除外を追加

構造柱・構造フレームの中から、型枠不要な鉄骨部材
（H形鋼・角形/円形鋼管・溝形鋼・山形鋼・CFT 等）を自動識別して除外する機能を追加。

### 識別ロジック (4 層フォールバック)
`Engine/SteelMemberDetector.cs`

| Layer | 内容 | 想定ヒット例 |
|---|---|---|
| L1 | `FamilyInstance.StructuralMaterialType == Steel` | 標準鋼材ファミリ |
| L2 | 断面形状分析 (`ExtrusionAnalyzer`) | CFT (中空)、H形鋼 (充実率<0.5) |
| L3 | 構造材マテリアル名 / `MaterialClass` に Steel/鋼/鉄/Metal | マテリアルだけ正しく設定された要素 |
| L4 | ファミリ名 / タイプ名のキーワードマッチ | 古い独自ファミリ、CFT-□400 等 |

- SRC柱は L2 で「中実・凸 (ratio≥0.5)」として保持される（型枠必要）
- CFT柱は L2 (中空モデリング時) または L4 ("CFT") で除外される
- 検出失敗時は保持側 (フェイルセーフ)
- 暗黙挙動として常に ON (`FormworkSettings.ExcludeSteelMembers = true`)、UI 露出なし

### データフロー
- `ElementCollector.CollectAndClassify` が `Targets` と `ExcludedSteel` の 2 リストを返す
- `FormworkResult.ExcludedSteelResults` に除外要素を記録（集計には含めない）
- `FormworkVisualizer.CreateExcludedSteelShapes` が元 Solid から DirectShape を作成
  - マーカー値 = `28Tools_Formwork_Steel`（通常マーカー `28Tools_Formwork` と区別）
  - 区分 = `鉄骨除外` (View Filter キー)
  - 部位 = `鉄骨(除外)`
  - 面積 = 0 ㎡

### 色分け
オレンジ系 `RGB(255, 145, 30)` を `_steelExcludedColor` で固定。
View Filter `型枠_鉄骨除外` でこの色を適用。

### 集計表からの除外
`ScheduleCreator` の Filter は `Equal MarkerValue` (`28Tools_Formwork`) のみ通すため、
鉄骨除外 DirectShape (`MarkerValueSteel`) は集計表に表示されない。

### クリーンアップ
`CleanupExistingFormworkShapes` は両方のマーカー値を持つ DirectShape を削除する
（再実行時の累積を防ぐ）。

### デバッグログ
`C:\temp\Formwork_debug.txt` に各要素の判定結果を出力:
```
---- Steel Member Detection ----
  [SteelExclude] E12345 Cat=構造柱 Name='H300x300' L=StructuralMaterialType reason=...
  [SteelKeep]    E12346 Cat=構造柱 Name='C700' reason=solid convex profile (areaRatio=0.987)
  Steel detection: target=42 excluded=8 kept=34
```

### 動作確認待ち
- 動作確認用モデルに RC 丸柱をモデリングして Layer 2 の誤検出がないか確認予定（ユーザー側）

---

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
