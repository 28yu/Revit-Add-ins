# 型枠数量算出アドイン 開発状況ハンドオフ

**最終更新**: 2026-05-14
**開発用 Revit**: 2022 (`dev-config.json`)

---

## 🔴 次セッション最優先課題（2026-05-14 時点）

**リンクモデルの壁スイープ面の型枠 DirectShape が作成されない問題**

詳細・仮説 A〜F・診断手順・検証コード例は `NEXT_SESSION_PROMPT.md` を参照。

### 当該セッション中の主な成果
- リンクモデル対応の全パイプライン実装（ElementSource / Registry / Transform / `28Tools_Formwork_ソース` / per-source 集計表 / ソース別小計サマリ）
- 7 カテゴリ目「屋根」を追加（勾配上向き面の自動天端化）
- 鉄骨階段の自動除外（`SteelStairDetector`）
- LGS 壁の自動除外（`LgsWallDetector` — 石膏ボード厚さ 9.5/12.5/15/21mm + コンクリート層なし）
- 壁の斜め天端控除（水平射影短辺 ≥ 30mm）
- 集計表の列幅を内容に応じて自動調整
- `MoveWallSweepsToExcluded` 除去（WallSweep を ElementResult として保持しスイープ外側面に型枠生成）

### 残課題
1. **リンクモデルの壁スイープ面の型枠が生成されない** ← 次セッションで取り組む

---

## 🎯 旧次セッションへの引き継ぎ事項（2026-05-11 終了時点）

### 動作確認済み・本番投入可能な機能
1. 6 カテゴリ（柱・梁・壁・床・基礎・階段）の型枠面積算出
2. 鉄骨部材の自動除外（4 層フォールバック判定）
3. デッキスラブの自動除外（タイプ名 "DS" 検出）
4. 壁スイープ・リビールの自動除外
5. 集計表「型枠数量集計」 + 動的合計サマリ「型枠数量集計_合計」
6. Excel 出力（CJK 対応）+ 解析 3D ビュー
7. **[53] 柱全面 Full Contact 誤判定の修正**: `Face.Project` null → offFace 扱い + `OverlapRatioMin=0.95` の 4 隅投影チェック
8. **[6] スラブ埋まり梁への型枠誤生成の修正**: partialSum ≥ 95% で DeductedContact に降格
9. **[7] 非矩形面の型枠が半透明問題の修正**: Clipper 失敗時に 3D Boolean Difference fallback
10. **[8] 直交梁端部型枠が省かれない問題の修正**: SmallAreaRatio=0.3 / OverlapRatioMinSmallA=0.25 で小さな a の閾値緩和
11. **[9] 構造基礎×スラブ重なり検出**: `cond3-project-null` も Stage 2 トリガーに追加
12. **[13] 壁 E4522721 の 290mm 帯型枠欠落の修正**: a が b より大きい場合の誤 Full Contact を 3D AABB 包含チェックで Partial に降格（2026-05-11）

### ⚠️ 未解決の課題（次セッションで取り組む）

#### [10] 構造基礎 E3733292 の梁接触面に「変な空白」

**症状**: 基礎側面 (face[3]) の部分接触は clipper-OK（面積値は正しい）だが、視覚的に空白/くり抜きが間違った位置にある。

**ログ（確認済み）**:
```
[FaceDiag] face[3] area=41.4411 partials=1 [ E3286572:a=10.7273 ] clipper-OK eff=30.7137

[Pair E3286572(Beam) x E3733292(Foundation)]
  f1->f3 dot=... overlap=0.500(ok=2/off=2) FULL_CONTACT  ← 梁端面が基礎側面に Full 受理

[Pair E3733292(Foundation) x E3286572(Beam)]
  f3->f1 dot=... cond2-area-ratio → Stage 2 → PARTIAL_CONTACT area=10.7273
```

**根本原因**:
- 逆方向（基礎 f3 を a、梁 f1 を b）の Stage 2 `EvaluatePartialBToA` で `ProjectBCornersToA` が ok=2/off=2 → null
- fallback の `EstimateUvOnAFromBSize` が梁の UV 寸法を基礎の UV 軸に当てはめるため、接触矩形の**位置は正しいが形状・向きがずれる**
- → Clipper が面の誤った位置を削り、視覚的空白が間違った場所に出る

**2026-05-08 で実装した改善（未確認）**:
`ProjectBCornersToARelaxed` を新規追加し、UvBoundsOnA 算出を 3 段階 fallback に変更:
1. strict: `ProjectBCornersToA` (ok≥3 必要)
2. **relaxed** (新規): 距離制限なし投影 → 境界外の隅は境界上の最近点 UV にスナップ
3. estimate: `EstimateUvOnAFromBSize` (UV 軸回転無視・最終手段)

デバッグログに `s2-uvOnA method=relaxed/strict/estimate` が出るので確認できる。

**次セッションでの確認手順**:
1. ビルドして Revit で再実行
2. `C:\temp\Formwork_debug.txt` で `s2-uvOnA method=` を確認 → `relaxed` になっているか
3. `[Clipper] subtract E3286572/f1 raw=... clipped=...` で UvBoundsOnA の位置が正しいか確認
   - 基礎側面 (2200×1750mm) 上で、梁端面の当たっている位置 (隅 or 辺側) に対応した UV 矩形か
4. 視覚的に空白が正しい位置に移ったか確認

**もし relaxed でも改善されない場合の次の手**:

方針 A: `Face.Project` を使わず、基礎面の **法線 + 原点からの平面方程式** で梁の隅を手動投影する。`PlanarFace` なら `face.ComputeNormal(UV.Zero)` と `face.Origin` (またはコーナー1点) から平面を構築し、梁の隅 3D 点をその平面に投影して UV を算出。API に依存せず確実に機能する。

方針 B: `[Clipper] subtract` ログの raw UV を分析し、実際の梁の位置と対応しているか検証。ずれのパターン（回転?スケール?オフセット?）を特定してから対処。

### 直近の修正コミット (順)
| 日時 | 内容 |
|---|---|
| 2026-05-11 | [13] 3D AABB 包含チェックで a>b の誤 Full Contact を Partial に降格 |
| 2026-05-11 | [13] PolygonCutter: face B の実ポリゴン形状でカット (L字面等に対応) |
| 2026-05-11 | [13] PolygonCutter: face B 法線逆向き時のループ自動反転 (needsReverse) |
| 2026-05-08 | [10] ProjectBCornersToARelaxed 追加 (3段階 fallback UvBoundsOnA) |
| 2026-05-08 | [9] cond3-project-null を Stage 2 トリガーに追加 |
| 2026-05-08 | [8] SmallAreaRatio=0.3 / OverlapRatioMinSmallA=0.25 で閾値動的切替 |
| 2026-05-08 | [7] TryBuildCarvedFaceSolid: Clipper 失敗時に 3D Boolean Difference fallback |
| 2026-05-08 | [6] partialSum ≥ 95% で DeductedContact 降格 |
| 2026-05-08 | [53] Face.Project null → offFace 扱い (exceptionCount 増やさない) |
| 2026-05-07 | 鉄骨除外 (4層判定) |
| 2026-05-07 | デッキスラブ除外 + 除外フィルタ既定非表示 |
| 2026-05-07 | Excel CJK 対応列幅 |
| 2026-05-07 | 集計表総合計を Excel と一致 |
| 2026-05-07 | 動的合計サマリ集計表 |
| 2026-05-07 | doc.Regenerate() で 3D ビュー描画問題解決 |
| 2026-05-07 | 壁スイープ除外 |
| 2026-05-07 | 部分接触の Clipper ベース面積計算 |
| 2026-05-07 | ElemDiag / FaceDiag 詳細診断ログ追加 |

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

## 2026-05-11 セッション: 壁 E4522721 の 290mm 帯型枠欠落を修正 [13]

### 問題の概要

天端切り欠き壁（高さ 1190.5mm）の上部 290mm 帯に型枠 DirectShape が表示されない。

**要素構成**:
- 壁 E4522721: 高さ 1190.5mm、上端に 110.5mm 高 × 75mm 深の切り欠き
  - face[1] (−Y 面): Z=−4.726〜−0.820 ft (全 1190.5mm の背面)
- 壁 E4485936: 隣接する低い壁、Z=−11.615〜−2.755 ft (上端が E4522721 の Z=−0.820 より低い)
- スラブ E4486237: Z=−2.756〜−1.772 ft

**症状の本当の原因**（最初の仮説は誤りだった）:
- face[0]（正面 +Y 面）ではなく **face[1]（背面 −Y 面）の誤 Full Contact 判定** が根本原因
- face[1] は E4485936 face[0] と `FULL_CONTACT` と判定 → `DeductedContact` → 型枠なし
- E4485936 の Z_max=−2.755 < face[1] の Z_max=−0.820（290mm = 0.951 ft はみ出し）

**誤判定のメカニズム**:
```
f1->f0 (E4522721 face[1] areaRatio=0.07 ≤ SmallAreaRatio=0.3)
→ 緩和閾値 OverlapRatioMinSmallA=0.25 を適用
→ okCount=2/4 (overlap=0.5) ≥ 0.25 → FULL_CONTACT ← 誤!
(E4522721 face[1] の Z_max が E4485936 face[0] の Z_max を 0.951ft 超えていた)
```

### 修正内容

#### 1. Stage 1 後の 3D AABB 包含チェック (`ContactFaceDetector.cs`)

Stage 1 が受理した後、a の 3D バウンディングボックスが b の 3D バウンディングボックスの範囲を超えて延びていれば Full Contact を拒否し、Partial へ降格する。

```csharp
if (stage1Accepted)
{
    var aBox = Compute3DBBox(a.Face);
    var bBox = Compute3DBBox(b.Face);
    if (aBox.min != null && bBox.min != null)
    {
        double tol = CoincidenceTolFeet * 2.0;
        bool aExtends =
            aBox.min.X < bBox.min.X - tol || aBox.max.X > bBox.max.X + tol ||
            aBox.min.Y < bBox.min.Y - tol || aBox.max.Y > bBox.max.Y + tol ||
            aBox.min.Z < bBox.min.Z - tol || aBox.max.Z > bBox.max.Z + tol;
        if (aExtends)
        {
            stage1Accepted = false;
            stage1Reason = "cond6-a-extends-beyond-b";
        }
    }
}
```

#### 2. `cond6-a-extends-beyond-b` を Stage 2 トリガーに追加

```csharp
else if (stage1Reason == "cond6-a-extends-beyond-b")
{
    stage2 = ComputePartialFromBBoxIntersect(a, b, out stage2Reason);
}
```

#### 3. 新規ヘルパー: `Compute3DBBox(Face face)`

UV bbox の 4 隅（+ 中点）を `face.Evaluate()` で 3D 点に変換し、3D AABB（min/max XYZ）を返す。

#### 4. 新規関数: `ComputePartialFromBBoxIntersect(FaceInfo a, FaceInfo b, out string reason)`

a と b の 3D AABB の交差範囲を計算し、それを a の UV 空間（`PlanarFace.XVector / YVector / Origin`）に投影して `ContactResult { Kind=Partial, UvBoundsOnA, FaceB }` を返す。

**PlanarFace UV 座標系の投影**:
```csharp
PlanarFace pf = (PlanarFace)a.Face;
XYZ origin = pf.Origin;
XYZ xVec = pf.XVector;
XYZ yVec = pf.YVector;
// 3D 点 → UV: u = dot(pt - origin, xVec), v = dot(pt - origin, yVec)
```

### 追加で実装: ポリゴンカッター (`FormworkVisualizer.cs`)

L 字形などの非矩形接触面に対して、UV-AABB 矩形よりも正確なカットを実現するため、
接触面 B の実際のエッジポリゴンを使ったカッターを実装した。

#### `BuildCutterFromContactFacePolygon`

```csharp
// face B のエッジループを取得
IList<CurveLoop> loops = contactFaceB.GetEdgesAsCurveLoops();

// face B の法線が押し出し方向 (face A の法線) と逆向きのとき → ループを反転
bool needsReverse = faceBNormal.DotProduct(dir) < 0;
if (needsReverse) { /* 各 CurveLoop を Reverse して向き修正 */ }

// 薄板 Solid として押し出し
Solid cutter = GeometryCreationUtilities.CreateExtrusionGeometry(
    loops, dir, prePad + thickness);
```

#### ⚠️ CurveLoop の向きの落とし穴

`CreateExtrusionGeometry` は押し出し方向から見て **CCW** のループを要求する。
face B の法線が面 A の法線（押し出し方向）と逆向きの場合、`GetEdgesAsCurveLoops()` は
face B の法線基準で CCW のループを返す → 押し出し方向から見ると CW になる。

**判定式**: `faceBNormal.DotProduct(faceANormal) < 0` → `needsReverse = true`

ループの反転方法:
```csharp
var reversed = new CurveLoop();
var curves = loop.ToList();
curves.Reverse();
foreach (var c in curves) reversed.Append(c.CreateReversed());
```

Revit は CW でも例外を投げずに裏返った（inside-out）Solid を生成することがあるため、
この問題は気づきにくい。ログに `needsReverse=True` が出たら要注意。

### 追加された定数

| 定数 | 値 | 追加理由 |
|---|---|---|
| `cond6-a-extends-beyond-b` | Stage 1 拒否理由文字列 | a が b の 3D AABB 外に延びている場合 |

### ログで確認できる新しいマーカー

```
[Pair E<a> x E<b>]
  f1->f0 ... cond6-a-extends-beyond-b → s2=bbox-intersect(area=X.XX)
  
[FormworkViz] polygon-cutter: loops=1 needsReverse=True volume=0.003
[BoolDiff] subtract volBefore=0.524 volAfter=0.312 delta=-0.212
```

---

## 2026-05-08 セッション: ContactFaceDetector 精度改善 [53][6][7][8][9][10]

### 解決した課題と実装内容

#### [53] 柱 E3280907 — 4 側面が全面 Full Contact に誤判定

**原因**: `ProjectBCornersToA` 内で `Face.Project` が null を返したとき、旧コードは `exceptionCount++` していた。`exceptionCount ≥ 2` で `return -1 (fail-safe → Full Contact 受理)` されるため、null を多く返す面（Join Geometry の複合面など）で誤って Full Contact 受理されていた。

**修正**: `proj == null || proj.UVPoint == null` のとき `offFaceCount++` に変更（null = 面外 = 重なっていない）。これで `overlapRatio = okCount/4.0` が正しく計算され、0.95 閾値で REJECTED → Stage 2 へフォールスルーする。

```csharp
// Before (ContactFaceDetector.cs ProjectBCornersToA)
IntersectionResult proj = null;
try { proj = a.Project(p); } catch { exceptionCount++; continue; }
if (proj == null || proj.UVPoint == null || ...) { exceptionCount++; continue; }

// After
try { proj = a.Project(p); } catch { /* swallow → off-face */ }
if (proj == null || proj.UVPoint == null || ...) { offFaceCount++; continue; }
```

**知見**: `Face.Project` は境界外の点で null を返すことがある（例外ではなく正常な off-face の意味）。exception と off-face を区別することが重要。

---

#### [6] スラブに埋まった梁に型枠_スラブが誤生成

**原因**: スラブ下面に梁上面が複数 Partial Contact → 各 pc.ContactArea はそれぞれ小さいが、合計すると面積の 95%+ を覆っていた。Clipper は全減算後の tiny な残り矩形から Solid を作ってしまっていた。

**修正**: `FormworkCalcEngine.ComputeAndSetEffectiveArea` 内で:
```csharp
double partialSum = fi.PartialContacts.Sum(pc => pc.ContactArea);
bool nearFullCoverage = partialSum >= fi.Area * 0.95;
bool tinyResidual = effectiveFeetSq < fi.Area * 0.01;
if (tinyResidual || nearFullCoverage) {
    fi.FaceType = FaceType.DeductedContact;  // 完全接触扱いに降格
    effectiveFeetSq = 0;
}
```

---

#### [7] 非矩形面（Clipper 失敗）の型枠が半透明

**原因**: Clipper は矩形面専用。非矩形面（TriangleFace・L字面など）は `clipper-fail:not-rectangular` になり、半透明 DirectShape になっていた。

**修正**: `FormworkVisualizer.TryBuildCarvedFaceSolid` を新規追加:
1. まず Clipper を試行
2. 失敗したら 3D Boolean Difference fallback:
   - 元面の Solid を薄板に押し出す
   - 各 PartialContact の UvBoundsOnA から「貫通する角材 Solid」を作成 (`BuildCutterFromUvRect`)
   - `BooleanOperationsUtils.ExecuteBooleanOperation(…Difference)` で削り取る

`BuildCutterFromUvRect`: `prePad=0.10`, `thickness=0.20` で外向きに貫通させることで、法線方向の符号不一致でも切削できる。

---

#### [8] 直交梁の端部型枠が省かれない

**原因**: 梁端面（小さい a）がホスト梁側面（大きい b）に接触するとき、b の Join Geometry の notch で一部の隅が off-face と判定 → `okCount=3, overlap=0.75 < 0.95` → REJECTED。

**修正**: 面積比が小さい（a/b ≤ 0.3）ときは閾値を緩和:
```csharp
private const double SmallAreaRatio = 0.3;
private const double OverlapRatioMinSmallA = 0.25;

double areaRatio = aArea / Math.Max(bArea, 1e-9);
double overlapThreshold = areaRatio <= SmallAreaRatio
    ? OverlapRatioMinSmallA   // = 0.25 (1/4 隅でOK)
    : OverlapRatioMin;        // = 0.95 (通常)
```

**根拠**: a が b より十分小さい場合、a の中心が b 上にあること（cond4 で確認済み）を優先し、notch による corner 外れは b 側の複雑形状のせいとして許容する。

---

#### [9] 構造基礎×スラブの重なり型枠（JG 未適用が原因と最初判明、その後 contact detection 漏れも確認）

**原因 (最終判明)**: スラブ下面が非常に広く、その中心が基礎上面の範囲外に位置する → `a.Face.Project(pA)` が null を返す → Stage 1 失敗理由 `cond3-project-null`。旧コードはこの理由では Stage 2 を試みなかった。

**修正**: Stage 2 トリガー条件に `cond3-project-null` を追加:
```csharp
if (stage1Reason == "cond2-area-ratio" ||
    stage1Reason == "cond5-overlap-insufficient" ||
    stage1Reason == "cond3-project-null")  // ← 追加
{
    stage2 = EvaluatePartialBToA(a, b, out stage2Reason);
}
```

**意味**: a の中心が b の範囲外 = a >> b という状況は `cond2-area-ratio` と意味的に同じ。大面 a に小面 b が部分接触しているケースとして Stage 2 で評価する。

---

#### EffectiveAreaM2 が DirectShape に伝播しない問題（[7] 修正時に発見）

**原因**: `FormworkCalcEngine.BuildElementResult` と `RecomputeFaces` が別々に `FaceInfo` を処理していたが、`EffectiveAreaM2` の計算（Clipper 呼び出し）が一方にしかなかった。`RecomputeFaces` 経由の場合は 0 のまま。

**修正**: `ComputeAndSetEffectiveArea(fi, out string clipperStatus)` ヘルパーに切り出し、両パスから呼び出す。このヘルパーが `fi.EffectiveAreaM2` を設定し、`FormworkVisualizer` がそれを DirectShape の厚み計算に使う。

---

### ContactFaceDetector の現在の定数一覧（2026-05-11 時点）

| 定数 | 値 | 意味 |
|---|---|---|
| `CoincidenceTolFeet` | 0.05 ft (≈15mm) | 接触面の許容距離 |
| `AreaRatioLimit` | 1.2 | Stage 1 の a/b 面積比上限 |
| `AntiParallelThreshold` | -0.90 | 法線の反平行判定 (cos ≤ -0.9) |
| `OverlapRatioMin` | 0.95 | Stage 1 の 4 隅重なり比閾値（通常） |
| `SmallAreaRatio` | 0.3 | a/b ≤ この値のとき「小面」扱い |
| `OverlapRatioMinSmallA` | 0.25 | 小面のときの緩和閾値 |

**⚠️ SmallAreaRatio / OverlapRatioMinSmallA の副作用と対策**:

[8] の修正（小面の緩和閾値）は意図通り動くが、副作用がある。
a が b より面積比で小さい（≤ 0.3）ときに閾値を緩和するため、
a が b の Z 範囲を超えて延びているケースでも Full Contact と誤判定してしまう
（okCount=2/4 → overlap=0.5 ≥ 0.25 でも通過）。

[13] の修正（3D AABB チェック）はこの副作用を打ち消す：Stage 1 受理後に
a が b の 3D AABB の外に延びていれば Full 受理を取り消し Partial 経路へ送る。

### EvaluatePartialBToA の UvBoundsOnA 3 段階 fallback（2026-05-08 実装）

```
1. ProjectBCornersToA(a, b, bbB)          → strict: ok≥3 隅が距離内
2. ProjectBCornersToARelaxed(a, b, bbB)   → relaxed: 距離制限なし (境界外→境界UV)
3. EstimateUvOnAFromBSize(bbA, bbB, uv)   → estimate: b の UV 寸法を a の軸に流用
```

ログで `s2-uvOnA method=strict/relaxed/estimate` として判定経路が確認できる。

### Clipper のデバッグログで UvBoundsOnA を確認する方法

`C:\temp\Formwork_debug.txt` で:
```
[Clipper] subtract E3286572/f1 raw=[U:1.23..2.45,V:0.11..1.34] clipped=[U:1.23..2.45,V:0.11..1.34]
```
- `raw` = UvBoundsOnA（クランプ前）
- `clipped` = 面の境界内にクランプ後

これを見て、実際に beam が接触している位置（基礎側面上の UV 座標）と一致しているか確認できる。
基礎 2200×1750mm なら UV 範囲は概ね [0..7.2, 0..5.7] ft。

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
