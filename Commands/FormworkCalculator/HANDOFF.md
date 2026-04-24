# 型枠数量算出アドイン 開発状況ハンドオフ

**最終更新**: 2026-04-24 (Phase 1: 部分接触検出 + 空間索引)
**ブランチ**: `claude/fix-formwork-contact-exclusion-6ZGi1` (最新)
**開発用 Revit**: 2022 (`dev-config.json`)

## 🆕 Phase 1: 部分接触 + 大規模対応 (2026-04-24)

[18]-3 の根本解決として「面の部分接触」を扱う仕組みを導入した。T字壁結合等で、主面 (大面積) の一部に他要素が当たっているケースを正しく面積控除できる。

### 追加された仕組み

1. **`PartialContact` データ構造** (`FaceClassifier.cs`)
   - 面が FormworkRequired のまま、部分的に覆われている領域を記録
   - `OtherElementId`, `OtherFaceIndex`, `ContactArea`, `UvBounds`

2. **`ContactFaceDetector` の2段階判定** (`ContactFaceDetector.cs`)
   - Stage 1: Full Contact (従来の area-ratio 1.5 以内) → 面全体を `DeductedContact`
   - Stage 2: Partial Contact (a > b × 1.5 + b が a 内部) → `PartialContacts` に追加、FormworkRequired は維持

3. **`SpatialGrid` 空間索引** (`SpatialGrid.cs`)
   - 全要素 BB を 3D 格子セルに登録、隣接セルだけ走査
   - セルサイズ = 全要素対角線長の中央値 × 2 (自動算出)
   - O(N²) → O(N) 相当になり 500+ 要素対応

4. **面積控除** (`FormworkCalcEngine.BuildElementResult`)
   - FormworkRequired 面の実効面積 = 元面積 - Σ(PartialContact.ContactArea)
   - `DeductedContactArea` に部分接触分も集計

5. **共有パラメータ** (`FormworkParameterManager.cs`)
   - `28Tools_Formwork_部分接触` (Text, "Yes"/"No") を追加
   - T字接合等で「一部消されている面」を識別可能

6. **半透明表示** (`FormworkVisualizer.cs`)
   - 部分接触がある面の DirectShape を 50% 透過で視覚的にヒント
   - `ApplyPartialContactOverride` でビュー単位オーバーライド

## 🆕 Phase 2: DirectShape 形状厳密化 (2026-04-24)

Phase 1 の面積控除に加え、DirectShape の形状自体も接触領域を切り抜いた厳密な矩形にする。

### アプローチ: 矩形ベース 2D 差分
1. 面 A の CurveLoop が軸平行矩形 (壁・梁・柱・スラブの主面で 90%+ のケース) か判定
2. 各 `PartialContact.UvBoundsOnA` の矩形を順次差分 (最大 4 サブ矩形に分解)
3. 各サブ矩形から薄板 Solid を構築 → 分割 DirectShape で生成

### ファイル追加/変更
- `Engine/PartialContactClipper.cs` — 矩形抽出・差分・Solid 構築
- `Engine/ContactFaceDetector.cs` — B の 4 隅を A に投影した `UvBoundsOnA` を算出
- `Output/FormworkVisualizer.cs` — Clipper 成功時は分割 DirectShape、失敗時は Phase 1 の半透明フォールバック

### 非対応ケース (フォールバック動作)
- 開口付き壁 (CurveLoop が穴を含む → 矩形判定で除外)
- カーブウォール (非対応要件)
- 非矩形の面 (4 頂点でない、曲線エッジを含む)
- 投影失敗 (UvBoundsOnA が null)

フォールバック時は Phase 1 の半透明 DirectShape で動作するため、リグレッションしない。

### 面積パラメータの配分
分割された DirectShape では、**最初の Solid** にのみ控除後面積を乗せ、残りは 0 にする。
これにより集計表の合計は従来通り要素毎の面積になる (重複加算を避ける)。

## 🔎 デバッグログ (2026-04-24 追加)

[18]-3 の原因特定のため、`ContactFaceDetector` の全判定を `C:\temp\Formwork_debug.txt` に出力するデバッグログ機構を導入した。

- **フラグ**: `FormworkSettings.EnableDebugLog` (デフォルト `true`)
- **集約クラス**: `Commands/FormworkCalculator/Engine/FormworkDebugLog.cs`
  - 実行ごとにファイルをクリア
  - 行数上限 200,000 行 (超えたら末尾に `... truncated` を追記して停止)
  - 内部で `lock` 排他制御
- **出力内容**:
  - `Pass 1` 完了時: 要素毎の面分類内訳 (`Req/Top/Bot/Con/BGL/Inc/Err`) とカテゴリ合計
  - `Pass 2`: BBox 重なりのある要素ペアごとに、反平行条件 (cond1) を通った面ペアの詳細
    - 法線内積 `dot`, 面積 `aArea/bArea`, 中心点 `pA`, 投影距離 `d`, `uv`, `bbB`
    - 採用/棄却 (`ACCEPTED` / `REJECTED(cond2-area-ratio|cond3-distance|cond4-uv-out-of-bounds|cond4-near-boundary|...)`)
  - `Pass 2` 完了時: 最終 FormworkRequired 面数, DeductedContact 面数, `contactChanges` 件数

**リリース時は `FormworkSettings.EnableDebugLog = false` に変更すること。** UI には露出していない。

**次のステップ**: ユーザーにテストケース (壁 T 字・梁柱取り合いで miss が出るモデル) で実行してもらい `Formwork_debug.txt` を共有してもらう → 失敗パターン (どの cond で落ちているか or ACCEPTED なのに視覚的には miss か) を特定して的確な修正に進む。

---

## 全体の進行状況

仕様書 (本アドインの元指示) の Phase 1〜5 は一通り実装済み。ユーザーのテストに応じて[1]〜[21]-1 の改善要望に対応し、ほとんどは解決した。

**唯一の未解決問題: [18]-3 要素間接触面の検出**

---

## 未解決問題の詳細: [18]-3

### 現象
壁の T 字結合や梁と柱の取り合い部で、接触している面にまだ型枠オブジェクト (DirectShape) が作られてしまう。本来はそこに型枠は不要。

### 試したアプローチ (全て一部ケースで miss あり)

| # | アプローチ | 採用期間 | 結果 |
|---|-----------|---------|------|
| 1 | 5 点 UV サンプリング + 距離 0.02ft | 初期 | 判定が厳しすぎて検出量が少ない |
| 2 | 中心点 + 面積比 + Face.Project.Distance 0.05ft | 第2世代 | Face.Project が境界外でも小さい Distance を返すため誤検出 |
| 3 | ReferenceIntersector (origin 外側 1mm) | 第3世代 | 密着面で隣要素の「反対側の面」にヒット → miss |
| 4 | ReferenceIntersector (origin 内側 10mm) | 第4世代 | 一部改善するも miss あり |
| 5 | ReferenceIntersector 両方向レイ | 第5世代 | 一部改善するも miss あり |
| 6 | ReferenceIntersector + 制限解除 rayView | 第6世代 | 一部改善するも miss あり |
| 7 | **(現状)** 幾何学的検査 + UV内部判定 + 面積比 | 最新 | まだ miss あり |

### 推定される原因
- Revit `Face.Project` は面境界外の点でも小さい perpendicular distance を返すが `UVPoint` は境界にクランプされる → UV内部判定で境界近傍を除外しているが Boolean Union 後の面分割で中心点が想定と違う位置にくるケースがある
- `ReferenceIntersector` は密着面 (Proximity ≈ 0) で挙動が不安定
- Join Geometry で結合された要素は面の形状が単純な予想と異なる場合がある

---

## 次セッションで推奨する方針

### 【優先】方針 A: デバッグログで原因特定
手を動かす前にまず **何がどこで落ちているかを可視化** する。

出力先: `C:\temp\Formwork_debug.txt`

ログすべき内容:
- Pass 1 完了時: 各要素の face 分類内訳 (Required/Top/Bottom/Contact/...)
- Pass 2 内: 全ての IsFaceCovered 判定について pass/fail の詳細
  - 要素A/B の Id, カテゴリ, 面 index, 中心点, 法線, 面積
  - 法線内積, Face.Project の Distance, UVPoint, UV bounds
  - **採用/棄却の理由** (条件 1-4 のどれで落ちたか)
- Pass 2 完了時: 最終 FormworkRequired 面数

ユーザーにテストしてもらいログを共有してもらうことで「**どの接触ペアが何の理由で落ちているか**」を特定できる。

### 方針 B: Boolean Union ベース (仕様書§3.1 の本来方針)
全要素の Solid をまとめて Union → 結合後の外面のみ抽出 → 内部接触面は自動消滅。
- 課題: 大量 Union は失敗しやすい、per-element attribution が失われる
- 対策: 近接要素グループに分割して Union 試行。attribution は Union 前に記録

### 方針 C: Boolean Intersection ベース
要素ペアの Intersection Solid を計算 → 体積 ≈ 0 かつ表面積 > 0 = 平面接触 → 対応する面を Contact 化

### 方針 D: 2D 多角形オーバーラップ
反平行+同一平面の面ペアで共通平面に両 CurveLoop を投影 → Sutherland-Hodgman で多角形交差 → 面 A 面積の 50% 以上が面 B に覆われていたら Contact

---

## 現在のコード構造

```
Commands/FormworkCalculator/
├── FormworkCalculatorCommand.cs       # エントリポイント
├── Models/
│   ├── FormworkSettings.cs            # UI設定
│   └── FormworkResult.cs              # 計算結果・面情報
├── Engine/
│   ├── ElementCollector.cs            # 要素収集 + レベル/パラメータ取得
│   ├── SolidUnionProcessor.cs         # Solid 取得 + Boolean Union
│   ├── FaceClassifier.cs              # 面分類 (Top/Bottom/Required)
│   ├── ContactFaceDetector.cs         # ★未解決: 接触面検出 (幾何検査版)
│   ├── OpeningProcessor.cs            # 開口部処理
│   ├── FormworkCalcEngine.cs          # メインエンジン (3-Pass)
│   ├── FormworkParameterManager.cs    # 共有パラメータ管理
│   └── FormworkFilterManager.cs       # View Filter 管理
├── Output/
│   ├── ExcelExporter.cs               # Excel 出力
│   ├── FormworkVisualizer.cs          # 3Dビュー + DirectShape
│   └── ScheduleCreator.cs             # ViewSchedule 作成
└── Views/
    ├── FormworkDialog.xaml            # メインダイアログ
    └── FormworkDialog.xaml.cs
```

### 3 Pass パイプライン (`FormworkCalcEngine.Run`)

```
Pass 1: 要素毎に Solid 取得 → FaceClassifier で分類
Pass 2: ContactFaceDetector で接触面を DeductedContact に変更 ← ★未解決
Pass 3: 開口加算 + ElementResult 作成 + Aggregate
```

### FaceType enum
- `FormworkRequired`: 型枠必要
- `DeductedTop`: 最上面（スラブは全上向き面）
- `DeductedBottom`: 最下面（基礎のみ、それ以外は FormworkRequired にコンバート）
- `DeductedContact`: 他要素との接触面 ← ★
- `DeductedBelowGL`: GL 以下
- `Inclined`: 傾斜面（現状未使用、全て FormworkRequired 扱い）
- `Error`: エラー

### 共有パラメータ (OST_GenericModel にバインド)
- `28Tools_FormworkMarker` (Text): DirectShape 識別マーカー = `"28Tools_Formwork"`
- `28Tools_Formwork_部位` (Text): 柱/梁/壁/スラブ/基礎/階段
- `28Tools_Formwork_レベル` (Text): 参照レベル名
- `28Tools_Formwork_区分` (Text): 色分けグループキー
- `28Tools_Formwork_面積` (Area): 面積 (㎡)

### ビュー・集計表
- 解析 3D ビュー名: `型枠分析` (日時なし・再実行で上書き)
- ViewSchedule 名: `型枠数量集計` (日時なし・再実行で上書き)
- Excel 初期名: `型枠数量集計.xlsx`
- 集計表のグループ化: レベル → 部位 → タイプ名 (ShowHeader/ShowFooter + IsItemized=false)

### 色分け
- `FormworkFilterManager` で View Filter ベース
- フィルタルール: `28Tools_Formwork_区分 == <groupKey>`
- 元躯体: RGB(94,94,94) + 20% 透過のオーバーライド

---

## ビルド・デプロイ関連

- AutoBuild がローカル Windows で常駐し、main への squash merge を検知して自動再ビルド→デプロイ
- ブランチ `claude/revit-formwork-addon-SRINI` に push → 自動で main に squash merge される
- ログ: `AutoBuild.log` / `AutoBuild_detail.log` (リポジトリ直下)
- 詳細は CLAUDE.md 参照
