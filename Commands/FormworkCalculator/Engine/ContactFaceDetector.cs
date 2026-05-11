using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 要素同士の接触面を検出する。2段階判定方式。
    ///
    /// 【Stage 1: Full Contact】(従来の動作)
    ///   面 a の中心が面 b の内部に投影でき、a の面積 ≤ b の面積 × 1.5
    ///   → 面 a 全体を DeductedContact に変更
    ///   (小面が大面に当たる典型: 梁端面→柱側面、壁端面→主壁側面)
    ///
    /// 【Stage 2: Partial Contact】(Phase 1 追加)
    ///   面 a の一部に小さな面 b が接している (a > b × 1.5 だが中心投影は成立)
    ///   → 面 a は FormworkRequired のまま、PartialContact リストに b の情報を追加
    ///   (T字壁結合: 主壁側面の一部に端壁端面が当たるケース)
    ///
    /// 【候補の絞り込み】SpatialGrid で隣接セル内の要素だけ走査 → 大規模対応。
    /// </summary>
    internal class ContactFaceDetector
    {
        internal class ElementFacesContext
        {
            public int ElementId;
            public CategoryGroup Category;
            public string CategoryName;
            public BoundingBoxXYZ BB;
            public List<FaceClassifier.FaceInfo> Faces = new List<FaceClassifier.FaceInfo>();
        }

        private const double CoincidenceTolFeet = 0.05;   // ≈ 15mm
        // Full Contact 判定の面積比上限。a の面積 ≤ b の面積 × このリミット のとき
        // Full Contact とみなす。値が大きいと「a がやや大きくても全面接触扱い」され
        // 過大控除になりやすいので、実際の接合パターン (柱面 vs 梁端面など) に
        // 対しては 1.2 程度の保守的値が妥当。
        private const double AreaRatioLimit = 1.2;
        private const double AntiParallelThreshold = -0.90;  // cos ≤ -0.9 (≈ 155°)
        // Stage 1 強化: a の 4 隅を b に投影した UV 矩形と b の UV 矩形の重なり比が
        // この値未満なら Full Contact を否定し Stage 2 評価へ流す。中心 1 点投影だけ
        // では co-planar な隣接面 (Join Geometry の影響等) を弾けないための補強。
        //
        // 0.95 の根拠: a がやや大きく、b が a に完全包含されるケースでも、
        // overlap ratio = bArea / aArea。AreaRatioLimit=1.2 のとき最悪ケースで
        // 1/1.2 ≈ 0.83 まで下がる。0.8 では co-planar 隣接面を弾けないので
        // 0.95 に厳格化し、わずかでも露出があれば Partial Contact 扱いにする。
        private const double OverlapRatioMin = 0.95;

        // a が b より十分小さい (a/b ≤ SmallAreaRatio) ときの緩和した閾値。
        // a が小さく b が大きい場合、a's corners は基本的に全部 b 上にあるはずだが、
        // b の Join Geometry の notch のせいで一部の corner が off-face と判定される
        // ことがある。0.25 (= 1/4 隅) で Full 受理することで「a の中心は b 上、
        // 距離 0」だが「b 側の複雑形状で corner check が信頼できない」ケースを救済。
        private const double SmallAreaRatio = 0.3;
        private const double OverlapRatioMinSmallA = 0.25;

        internal static void RefineContactFaces(List<ElementFacesContext> contexts)
        {
            int n = contexts.Count;
            int contactChanges = 0;
            int partialContactAdded = 0;
            int pairsChecked = 0;
            int pairsBBoxOverlap = 0;

            FormworkDebugLog.Section($"Pass 2: Contact Face Detection (elements={n})");

            // 空間索引で候補絞り込み
            var grid = SpatialGrid.Build(contexts, CoincidenceTolFeet);
            FormworkDebugLog.Log($"SpatialGrid cell size: {grid.CellSize:F3} ft");

            var ctxById = new Dictionary<int, ElementFacesContext>();
            foreach (var c in contexts) ctxById[c.ElementId] = c;

            foreach (var ci in contexts)
            {
                if (ci.BB == null) continue;

                foreach (var otherId in grid.GetCandidates(ci.ElementId))
                {
                    if (!ctxById.TryGetValue(otherId, out var cj)) continue;
                    pairsChecked++;
                    pairsBBoxOverlap++;

                    bool pairLogged = false;

                    for (int fiIdx = 0; fiIdx < ci.Faces.Count; fiIdx++)
                    {
                        var fi = ci.Faces[fiIdx];

                        // FormworkRequired 面のみが対象
                        if (fi.FaceType != FaceType.FormworkRequired) continue;

                        // 既に Full Contact 化された面は以降スキップ
                        bool fullyContacted = false;

                        for (int fjIdx = 0; fjIdx < cj.Faces.Count; fjIdx++)
                        {
                            if (fullyContacted) break;
                            var fj = cj.Faces[fjIdx];
                            var result = EvaluateContact(
                                ci, fi, fiIdx,
                                cj, fj, fjIdx,
                                ref pairLogged);

                            if (result.Kind == ContactKind.Full)
                            {
                                fi.FaceType = FaceType.DeductedContact;
                                contactChanges++;
                                fullyContacted = true;
                            }
                            else if (result.Kind == ContactKind.Partial)
                            {
                                // 重複を防ぐ (同じ相手要素・同じ面を2回加算しない)
                                bool dup = false;
                                foreach (var pc in fi.PartialContacts)
                                {
                                    if (pc.OtherElementId == cj.ElementId &&
                                        pc.OtherFaceIndex == fjIdx) { dup = true; break; }
                                }
                                if (!dup)
                                {
                                    fi.PartialContacts.Add(new FaceClassifier.PartialContact
                                    {
                                        OtherElementId = cj.ElementId,
                                        OtherFaceIndex = fjIdx,
                                        ContactArea = result.ContactArea,
                                        UvBounds = result.UvBounds,
                                        UvBoundsOnA = result.UvBoundsOnA,
                                    });
                                    partialContactAdded++;
                                }
                            }
                        }
                    }
                }
            }

            FormworkDebugLog.Section("Pass 2 Summary");
            FormworkDebugLog.Log($"pairs checked (grid):        {pairsChecked}");
            FormworkDebugLog.Log($"pairs with BB overlap:       {pairsBBoxOverlap}");
            FormworkDebugLog.Log($"full contact changes:        {contactChanges}");
            FormworkDebugLog.Log($"partial contacts added:      {partialContactAdded}");
            FormworkDebugLog.Flush();
        }

        private enum ContactKind { None, Full, Partial }

        private struct ContactResult
        {
            public ContactKind Kind;
            public double ContactArea;
            public BoundingBoxUV UvBounds;       // B 自身の UV (debug)
            public BoundingBoxUV UvBoundsOnA;    // A 面上での B の投影 UV (Phase 2 用)
        }

        /// <summary>
        /// 面 a に対する面 b の接触を評価する。None / Full / Partial のいずれかを返す。
        /// </summary>
        private static ContactResult EvaluateContact(
            ElementFacesContext ca, FaceClassifier.FaceInfo a, int aIdx,
            ElementFacesContext cb, FaceClassifier.FaceInfo b, int bIdx,
            ref bool pairHeaderLogged)
        {
            var none = new ContactResult { Kind = ContactKind.None };
            if (a?.Face == null || b?.Face == null) return none;
            if (a.Normal == null || b.Normal == null) return none;

            bool log = FormworkDebugLog.Enabled;

            // --- 条件 1: 反平行 ---
            double dot = a.Normal.DotProduct(b.Normal);
            bool cond1 = dot <= AntiParallelThreshold;
            if (!cond1) return none;

            // --- 面積取得 ---
            double aArea = 0, bArea = 0;
            try { aArea = a.Face.Area; bArea = b.Face.Area; } catch { }
            bool cond2Area = aArea > 1e-6 && bArea > 1e-6;

            // --- 投影と距離 ---
            BoundingBoxUV bbA = null;
            try { bbA = a.Face.GetBoundingBox(); } catch { }
            XYZ pA = null;
            if (bbA != null) { try { pA = a.Face.Evaluate((bbA.Min + bbA.Max) * 0.5); } catch { } }

            IntersectionResult proj = null;
            if (pA != null) { try { proj = b.Face.Project(pA); } catch { } }
            bool cond3Dist = proj != null && proj.Distance <= CoincidenceTolFeet;

            UV uv = proj?.UVPoint;
            BoundingBoxUV bbB = null;
            try { bbB = b.Face.GetBoundingBox(); } catch { }

            bool cond4UVInside = false;
            bool cond4NotNearBoundary = false;
            if (uv != null && bbB != null)
            {
                double marginU = (bbB.Max.U - bbB.Min.U) * 0.01;
                double marginV = (bbB.Max.V - bbB.Min.V) * 0.01;

                cond4UVInside =
                    uv.U >= bbB.Min.U - 1e-6 && uv.U <= bbB.Max.U + 1e-6 &&
                    uv.V >= bbB.Min.V - 1e-6 && uv.V <= bbB.Max.V + 1e-6;

                bool nearBoundary =
                    uv.U < bbB.Min.U + marginU || uv.U > bbB.Max.U - marginU ||
                    uv.V < bbB.Min.V + marginV || uv.V > bbB.Max.V - marginV;

                cond4NotNearBoundary = !nearBoundary ||
                    (proj != null && proj.Distance <= CoincidenceTolFeet * 0.5);
            }

            // --- Stage 1: Full Contact 判定 (従来ロジック a→b の向き) ---
            bool stage1Accepted = false;
            string stage1Reason = null;
            double overlapRatio = -1;
            string overlapFailReason = null;
            if (!cond2Area) stage1Reason = "cond2-area-zero";
            else if (aArea > bArea * AreaRatioLimit) stage1Reason = "cond2-area-ratio";
            else if (bbA == null) stage1Reason = "cond3-bbA-null";
            else if (pA == null) stage1Reason = "cond3-pA-null";
            else if (proj == null) stage1Reason = "cond3-project-null";
            else if (!cond3Dist) stage1Reason = "cond3-distance";
            else if (uv == null) stage1Reason = "cond4-uv-null";
            else if (bbB == null) stage1Reason = "cond4-bbB-null";
            else if (!cond4UVInside) stage1Reason = "cond4-uv-out-of-bounds";
            else if (!cond4NotNearBoundary) stage1Reason = "cond4-near-boundary";
            else
            {
                // Stage 1 強化: a の 4 隅を b に投影した UV 矩形と b の UV 矩形の
                // 重なり比をチェック。中心 1 点投影だけでは co-planar な隣接面を弾けない。
                overlapRatio = EstimateOverlapRatioAonB(
                    a.Face, b.Face, bbA, bbB, out overlapFailReason);

                // a が b より十分小さい場合 (a/b ≤ 0.3) は閾値を緩和する。
                // 理由: a が小さく b が大きい場合、a's corners は基本的にすべて b 内に
                // 入るはずだが、b の Join Geometry による notch / 穴のせいで一部 corner
                // が off-face と判定されることがある。これは a の overhang ではなく
                // b 側の複雑形状が原因なので、閾値を下げて 1 隅でも乗っていれば Full
                // 扱いにする (a の中心は cond4 で b 上にあると確認済み)。
                //
                // 例: 直交梁の端面 (17.58 ft²) がホスト梁側面 (163.95 ft²) に接触する
                //    ケース。a/b = 0.107、okCount=3/4 (1 隅が b の notch にハマる)。
                //    旧 0.95 閾値だと REJECTED → Stage 2 も失敗 → 端面型枠が残る不具合。
                double areaRatio = aArea / Math.Max(bArea, 1e-9);
                double overlapThreshold = areaRatio <= SmallAreaRatio
                    ? OverlapRatioMinSmallA
                    : OverlapRatioMin;

                if (overlapRatio < 0)
                {
                    // 投影失敗: 既存ロジック (中心投影 + 面積比) を信頼して Full 受理
                    stage1Accepted = true;
                }
                else if (overlapRatio < overlapThreshold)
                {
                    stage1Reason = "cond5-overlap-insufficient";
                }
                else
                {
                    stage1Accepted = true;
                }
            }

            // --- Stage 2: Partial Contact 判定 (a >> b の場合) ---
            //   大面 a に小面 b が部分的に接触しているケース
            //   判定: b の中心を a に投影して、a 内部にあり距離≈0ならば部分接触
            ContactResult stage2 = none;
            string stage2Reason = null;
            if (!stage1Accepted)
            {
                // Stage 1 が以下の理由で否定された場合のみ Stage 2 (Partial Contact) を試す:
                // - cond2-area-ratio: a >> b で面積比制約に引っかかった
                // - cond5-overlap-insufficient: a の corner が b 上に乗り切らない
                // - cond3-project-null: a の中心を b に投影できない (a が b より遥かに広く、
                //   a の中心が b の範囲外に位置するケース。意味的には cond2-area-ratio と
                //   同類。例: slab 下面 (15.2m²) と foundation 上面 (12.9m²) で slab 中心が
                //   foundation 範囲外に来て Project が null を返す)
                if (stage1Reason == "cond2-area-ratio" ||
                    stage1Reason == "cond5-overlap-insufficient" ||
                    stage1Reason == "cond3-project-null")
                {
                    stage2 = EvaluatePartialBToA(a, b, out stage2Reason);
                }
                else
                {
                    stage2Reason = "not-evaluated";
                }
            }

            // --- ログ出力 ---
            if (log)
            {
                if (!pairHeaderLogged)
                {
                    FormworkDebugLog.Log(
                        $"[Pair E{ca.ElementId}({ca.Category}) x E{cb.ElementId}({cb.Category})]");
                    pairHeaderLogged = true;
                }

                var sb = new System.Text.StringBuilder();
                sb.Append("  f").Append(aIdx).Append("->f").Append(bIdx);
                sb.Append(" dot=").Append(Fmt(dot, 3));
                sb.Append(" aArea=").Append(Fmt(aArea, 4));
                sb.Append(" bArea=").Append(Fmt(bArea, 4));
                sb.Append(" d=").Append(proj != null ? Fmt(proj.Distance, 4) : "-");
                sb.Append(" uv=").Append(FmtUV(uv));
                sb.Append(" overlap=");
                if (overlapRatio >= 0) sb.Append(Fmt(overlapRatio, 3));
                else sb.Append("-");
                if (overlapFailReason != null)
                    sb.Append("(").Append(overlapFailReason).Append(")");
                if (stage1Accepted)
                {
                    sb.Append(" FULL_CONTACT");
                }
                else if (stage2.Kind == ContactKind.Partial)
                {
                    sb.Append(" PARTIAL_CONTACT area=").Append(Fmt(stage2.ContactArea, 4));
                }
                else
                {
                    sb.Append(" REJECTED(s1=").Append(stage1Reason ?? "-")
                      .Append(",s2=").Append(stage2Reason ?? "-").Append(")");
                }
                FormworkDebugLog.Log(sb.ToString());
            }

            if (stage1Accepted) return new ContactResult { Kind = ContactKind.Full };
            if (stage2.Kind == ContactKind.Partial) return stage2;
            return none;
        }

        /// <summary>
        /// 面 b の中心を面 a に投影して、a 内部に b が「一部として」収まるかを評価する。
        /// 成立するなら面 b の面積を Partial Contact 面積として返す。
        /// (a >> b の前提で、b 全体が a に覆われている想定)
        ///
        /// Phase 2 用に、b の 4 隅を a に投影した UV 矩形 (UvBoundsOnA) も算出する。
        /// </summary>
        private static ContactResult EvaluatePartialBToA(
            FaceClassifier.FaceInfo a, FaceClassifier.FaceInfo b, out string reason)
        {
            reason = null;
            var none = new ContactResult { Kind = ContactKind.None };

            if (a?.Face == null || b?.Face == null) { reason = "s2-face-null"; return none; }

            BoundingBoxUV bbB = null;
            try { bbB = b.Face.GetBoundingBox(); } catch { }
            if (bbB == null) { reason = "s2-bbB-null"; return none; }

            XYZ pB = null;
            try { pB = b.Face.Evaluate((bbB.Min + bbB.Max) * 0.5); } catch { }
            if (pB == null) { reason = "s2-pB-null"; return none; }

            // b の中心を a に投影して、距離チェックと UV 取得を行う。
            // a.Face.Project が null を返す場合 (壁-壁 T字接合等で pB が a の UV 境界
            // 近傍に位置するとき Revit が null を返す既知の挙動) は、a が PlanarFace なら
            // 平面方程式で補完する。
            IntersectionResult proj = null;
            try { proj = a.Face.Project(pB); } catch { }

            double projDist;
            UV uv;
            if (proj != null)
            {
                projDist = proj.Distance;
                uv = proj.UVPoint;
            }
            else
            {
                var pfa = a.Face as PlanarFace;
                if (pfa == null) { reason = "s2-project-null"; return none; }

                XYZ xVec = pfa.XVector;
                XYZ yVec = pfa.YVector;
                if (xVec == null || yVec == null) { reason = "s2-project-null"; return none; }

                XYZ normal = xVec.CrossProduct(yVec);
                if (normal.GetLength() < 1e-9) { reason = "s2-project-null"; return none; }
                normal = normal.Normalize();

                XYZ d3d = pB - pfa.Origin;
                projDist = Math.Abs(d3d.DotProduct(normal));

                XYZ xN = xVec.Normalize();
                XYZ yN = yVec.Normalize();
                uv = new UV(d3d.DotProduct(xN), d3d.DotProduct(yN));

                if (FormworkDebugLog.Enabled)
                    FormworkDebugLog.Log($"    s2-via-plane planeDist={projDist:F6} uv={FmtUV(uv)}");
            }

            if (projDist > CoincidenceTolFeet) { reason = "s2-distance"; return none; }

            BoundingBoxUV bbA = null;
            try { bbA = a.Face.GetBoundingBox(); } catch { }
            if (uv == null || bbA == null) { reason = "s2-uv-null"; return none; }

            // b の中心投影が a の UV 内部か? (境界付近でもOK、全体が覆われている前提)
            double marginU = (bbA.Max.U - bbA.Min.U) * 0.001;
            double marginV = (bbA.Max.V - bbA.Min.V) * 0.001;
            bool uvInside =
                uv.U >= bbA.Min.U - marginU && uv.U <= bbA.Max.U + marginU &&
                uv.V >= bbA.Min.V - marginV && uv.V <= bbA.Max.V + marginV;

            // 中心が a の UV 外でも、b の 4 隅投影が a と重なる場合は partial contact として扱う。
            // 典型ケース: a が大面 (e.g. 床長辺 5385x300mm)、b がオフセットした小面
            // (e.g. 壁背面 825x1191mm, 高さ範囲が a と部分的にしか重ならない)。
            // 中心点ベースの判定だけだと「a と b が同じ平面上にあるが中心同士が離れている」
            // 部分接触ケースを取りこぼす。
            bool centerOutsideButCornersMayOverlap = !uvInside && projDist <= CoincidenceTolFeet;

            if (!uvInside && !centerOutsideButCornersMayOverlap)
            {
                reason = "s2-uv-out-of-bounds";
                return none;
            }

            double bArea = 0;
            try { bArea = b.Face.Area; } catch { }
            if (bArea <= 1e-6) { reason = "s2-b-area-zero"; return none; }

            // Phase 2: b の 4 隅を a に投影した UV 矩形を算出
            //
            // a が PlanarFace の場合は ViaPlane (平面方程式 + XVector/YVector) を最優先で使う。
            // Face.Project は b の隅が a の polygon 外 (L字形 face の notch 部等) に
            // 落ちると境界にスナップして UV が拡張されてしまう不具合がある:
            //   例: face[0] (L字, V max=3.906) に対し b/f6 (V max=3.815) を strict 投影すると
            //       V=3.906 にスナップ → 全高ストリップが cutter になり、上部の型枠まで
            //       誤って carve される。
            // ViaPlane は境界スナップしないので正確な UV 範囲が得られる。
            BoundingBoxUV uvOnA = null;
            string uvOnAMethod = null;

            if (a.Face is PlanarFace)
            {
                uvOnA = ProjectBCornersToAViaPlane(a.Face, b.Face, bbB);
                if (uvOnA != null) uvOnAMethod = "via-plane";
            }

            // Fallback 1: ViaPlane が使えない/失敗した場合は strict (Face.Project) を試す。
            if (uvOnA == null)
            {
                uvOnA = ProjectBCornersToA(a.Face, b.Face, bbB);
                if (uvOnA != null) uvOnAMethod = "strict";
            }

            // Fallback 2: 非平面 (CylindricalFace 等) の場合、距離制限なしで Face.Project する。
            // b の隅が a の境界外にはみ出していると a.Project は境界上の最近点 UV にスナップする。
            // PartialContactClipper でクランプされる前提のため位置精度は落ちるが via-plane が
            // 使えない場合のフォールバック。
            if (uvOnA == null)
            {
                uvOnA = ProjectBCornersToARelaxed(a.Face, b.Face, bbB);
                if (uvOnA != null) uvOnAMethod = "relaxed";
            }

            // Fallback 3: 寛容な投影も null の場合 (a.Project が全て null 等)、
            // b の UV 寸法を使って `uv` 中心の矩形を構築する。
            // a と b の UV 軸の相対回転が大きい場合はずれが生じるが最後の手段として使用。
            if (uvOnA == null)
            {
                uvOnA = EstimateUvOnAFromBSize(bbA, bbB, uv);
                if (uvOnA != null) uvOnAMethod = "estimate";
            }

            if (FormworkDebugLog.Enabled && uvOnAMethod != null)
            {
                FormworkDebugLog.Log($"    s2-uvOnA method={uvOnAMethod} bounds={FmtUVBounds(uvOnA)}");
            }

            // uvOnA を face A の UV 範囲でクランプ。
            // B の投影コーナーが A の外にはみ出すと bounding box が膨れる（area > faceA.Area も起こりうる）。
            // A の UV 範囲に制限して Clipper に正確な接触 UV 矩形を渡す。
            // (ok=1/off=3 などのケース: 3 コーナーが A 外 → 膨れた bbox → Clipper 誤判定 → demoted-to-contact)
            if (uvOnA != null)
            {
                double clampUMin = Math.Max(uvOnA.Min.U, bbA.Min.U);
                double clampUMax = Math.Min(uvOnA.Max.U, bbA.Max.U);
                double clampVMin = Math.Max(uvOnA.Min.V, bbA.Min.V);
                double clampVMax = Math.Min(uvOnA.Max.V, bbA.Max.V);
                if (clampUMax > clampUMin && clampVMax > clampVMin)
                {
                    if (FormworkDebugLog.Enabled && (
                        Math.Abs(clampUMin - uvOnA.Min.U) > 0.001 ||
                        Math.Abs(clampUMax - uvOnA.Max.U) > 0.001 ||
                        Math.Abs(clampVMin - uvOnA.Min.V) > 0.001 ||
                        Math.Abs(clampVMax - uvOnA.Max.V) > 0.001))
                    {
                        FormworkDebugLog.Log(
                            $"    s2-uvOnA clamped to faceA bounds={FmtUVBounds(uvOnA)} → [{clampUMin:F3}..{clampUMax:F3},{clampVMin:F3}..{clampVMax:F3}]");
                    }
                    uvOnA = new BoundingBoxUV(clampUMin, clampVMin, clampUMax, clampVMax);
                }
                else
                {
                    // クランプ後に UV 範囲が消えた = 実際には A の UV 空間と重ならない
                    FormworkDebugLog.Log($"    s2-uvOnA clamp-eliminated (no overlap with faceA UV range)");
                    uvOnA = null;
                }
            }

            // ContactArea の推定:
            // face A が L 字等の非矩形の場合、UV bbox 面積は実ポリゴン面積より大きい
            // (例: face[5] 実面積=38.66ft² vs UV bbox=84.92ft², 充填率=45.5%)。
            // クランプ済み uvOnA がそのまま使えるとは限らない (uvOnA が UV bbox を
            // 超えなくても、実ポリゴン外の領域を含むため過大評価される)。
            //
            // 解決策: 面ポリゴンが UV bbox 内に均一分布すると仮定した期待接触面積。
            //   contactArea ≈ (uvOnA面積 / face A UV bbox面積) × face A 実面積
            //
            // [12] 梁 face[5] (L字, 38.66ft²) x スラブ面 (46.26ft²) のケースで
            // 旧実装は bArea=46.26 を使い partialSum>face*0.95 で誤 demoted-to-contact。
            // 修正後は 46.26/84.92×38.66 = 21.07 となり demoted されず、上部に型枠が生成される。
            // 中心が a の UV 外でクランプ後 uvOnA が null (= corners も a と重ならない)
            // 場合は真に接触なし。reject する (full bArea を引いて過剰控除するのを防ぐ)。
            if (!uvInside && uvOnA == null)
            {
                reason = "s2-uv-out-of-bounds";
                return none;
            }

            double aFaceArea = 0;
            try { aFaceArea = a.Face.Area; } catch { }
            double faceUvBboxArea = (bbA.Max.U - bbA.Min.U) * (bbA.Max.V - bbA.Min.V);
            double contactAreaEstimate = bArea;
            if (uvOnA != null && faceUvBboxArea > 1e-9 && aFaceArea > 1e-9)
            {
                double uvRectArea = (uvOnA.Max.U - uvOnA.Min.U) * (uvOnA.Max.V - uvOnA.Min.V);
                contactAreaEstimate = (uvRectArea / faceUvBboxArea) * aFaceArea;
                contactAreaEstimate = Math.Min(contactAreaEstimate, bArea);
            }

            return new ContactResult
            {
                Kind = ContactKind.Partial,
                ContactArea = contactAreaEstimate,
                UvBounds = bbB,
                UvBoundsOnA = uvOnA,
            };
        }

        /// <summary>
        /// 4 隅投影が失敗したときの保守的な UvBoundsOnA 推定。
        /// b の UV 矩形寸法を a の UV 系に流用し、`uv` を中心とする矩形を返す。
        /// a の境界外にはみ出る分は呼び出し側 (PartialContactClipper) でクランプされる。
        /// </summary>
        private static BoundingBoxUV EstimateUvOnAFromBSize(
            BoundingBoxUV bbA, BoundingBoxUV bbB, UV centerOnA)
        {
            if (bbA == null || bbB == null || centerOnA == null) return null;
            double halfU = (bbB.Max.U - bbB.Min.U) * 0.5;
            double halfV = (bbB.Max.V - bbB.Min.V) * 0.5;
            if (halfU <= 1e-6 || halfV <= 1e-6) return null;
            return new BoundingBoxUV(
                centerOnA.U - halfU, centerOnA.V - halfV,
                centerOnA.U + halfU, centerOnA.V + halfV);
        }

        /// <summary>
        /// Phase 2 Fallback 1 用: a が PlanarFace の場合、XVector/YVector を使って
        /// b の 4 隅を a の UV 系に「平面方程式から直接」変換する。
        ///
        /// PlanarFace の UV パラメタリゼーション:
        ///   p = Origin + u * XVector + v * YVector
        /// 逆変換 (XVector, YVector が単位ベクトルの場合):
        ///   u = dot(p - Origin, XVector)
        ///   v = dot(p - Origin, YVector)
        ///
        /// Face.Project と異なり境界外でも正確な UV を返す (境界スナップなし)。
        /// a が PlanarFace でない場合は null を返し、relaxed フォールバックに任せる。
        /// </summary>
        private static BoundingBoxUV ProjectBCornersToAViaPlane(Face a, Face b, BoundingBoxUV bbB)
        {
            if (bbB == null) return null;
            var pf = a as PlanarFace;
            if (pf == null) return null;

            XYZ origin = pf.Origin;
            XYZ xVec = pf.XVector;
            XYZ yVec = pf.YVector;
            if (origin == null || xVec == null || yVec == null) return null;

            double xLen2 = xVec.DotProduct(xVec);
            double yLen2 = yVec.DotProduct(yVec);
            if (xLen2 < 1e-12 || yLen2 < 1e-12) return null;

            // 通常 Revit の PlanarFace.XVector/YVector は unit vector だが、
            // 念のため正規化しておく。
            XYZ xN = xVec.Normalize();
            XYZ yN = yVec.Normalize();

            // b の actual polygon の境界エッジから 3D 頂点を収集する。
            // bbB の 4 隅を直接使うと b が L 字形等の場合に bbox の notch 内点も含めて
            // しまい、a 上の UV 範囲が過大になる (典型例: 隣接壁 f6 が wall top notch
            // と同じ場所に notch を持つケース)。
            // 実エッジを使えば polygon の実際の範囲が反映される。
            var pts = new List<XYZ>();
            try
            {
                var loops = b.GetEdgesAsCurveLoops();
                if (loops != null && loops.Count > 0)
                {
                    foreach (var loop in loops)
                    {
                        foreach (var curve in loop)
                        {
                            if (curve == null) continue;
                            // Curve を 8 分割で tessellate (Line は端点 2 点で十分だが
                            // Arc 等のために統一)
                            for (int i = 0; i <= 8; i++)
                            {
                                double t = i / 8.0;
                                try
                                {
                                    var p = curve.Evaluate(t, true);
                                    if (p != null) pts.Add(p);
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            catch { }

            // fallback: エッジ取得失敗時は bbB の 4 隅を使う
            if (pts.Count < 3)
            {
                pts.Clear();
                var corners = new UV[]
                {
                    new UV(bbB.Min.U, bbB.Min.V),
                    new UV(bbB.Max.U, bbB.Min.V),
                    new UV(bbB.Max.U, bbB.Max.V),
                    new UV(bbB.Min.U, bbB.Max.V),
                };
                foreach (var c in corners)
                {
                    try { var p = b.Evaluate(c); if (p != null) pts.Add(p); } catch { }
                }
            }

            double uMin = double.MaxValue, vMin = double.MaxValue;
            double uMax = double.MinValue, vMax = double.MinValue;
            int successCount = 0;

            foreach (var p in pts)
            {
                XYZ d = p - origin;
                double u = d.DotProduct(xN);
                double v = d.DotProduct(yN);

                if (u < uMin) uMin = u;
                if (v < vMin) vMin = v;
                if (u > uMax) uMax = u;
                if (v > vMax) vMax = v;
                successCount++;
            }

            if (successCount < 3) return null;
            if (uMax <= uMin || vMax <= vMin) return null;

            return new BoundingBoxUV(uMin, vMin, uMax, vMax);
        }

        /// <summary>
        /// Phase 2 Fallback 2 用: b の 4 隅を a の面に「寛容な投影」で投影し、UV AABB を返す。
        /// ProjectBCornersToA と異なり Distance の上限チェックを行わない。
        /// b の隅が a の境界外にある場合、a.Project は境界上の最近点 UV を返すため、
        /// 接触範囲の UV AABB を面境界にスナップした形で推定できる。
        /// </summary>
        private static BoundingBoxUV ProjectBCornersToARelaxed(Face a, Face b, BoundingBoxUV bbB)
        {
            if (bbB == null) return null;
            var corners = new UV[]
            {
                new UV(bbB.Min.U, bbB.Min.V),
                new UV(bbB.Max.U, bbB.Min.V),
                new UV(bbB.Max.U, bbB.Max.V),
                new UV(bbB.Min.U, bbB.Max.V),
            };

            double uMin = double.MaxValue, vMin = double.MaxValue;
            double uMax = double.MinValue, vMax = double.MinValue;
            int successCount = 0;

            foreach (var c in corners)
            {
                XYZ p;
                try { p = b.Evaluate(c); } catch { continue; }
                if (p == null) continue;

                // 距離制限なし: face 境界外の点は境界上の最近点 UV に収束する。
                // b は a に接触していることが確認済みなので距離チェックは不要。
                IntersectionResult proj = null;
                try { proj = a.Project(p); } catch { }
                if (proj == null || proj.UVPoint == null) continue;

                var uv = proj.UVPoint;
                if (uv.U < uMin) uMin = uv.U;
                if (uv.V < vMin) vMin = uv.V;
                if (uv.U > uMax) uMax = uv.U;
                if (uv.V > vMax) vMax = uv.V;
                successCount++;
            }

            if (successCount < 2) return null;
            if (uMax <= uMin || vMax <= vMin) return null;

            return new BoundingBoxUV(uMin, vMin, uMax, vMax);
        }

        /// <summary>
        /// Stage 1 強化用: 面 a の 4 隅を面 b に投影し、その AABB と b の UV 矩形の
        /// 交差面積を a の投影 AABB 面積で割った重なり比を返す。
        ///
        /// 1.0 = a 全体が b の UV 内、0.0 = 重なりなし、-1 = 投影失敗 (判定不能)
        ///
        /// 平面同士の場合は UV ≈ world 座標 (ft) なので世界面積比とほぼ等価。
        /// 中心 1 点投影だけでは弾けない co-planar 隣接面を識別するために使用。
        ///
        /// `failReason` には -1 を返した理由が入る (投影失敗の切り分け用)。
        /// </summary>
        private static double EstimateOverlapRatioAonB(
            Face a, Face b, BoundingBoxUV bbA, BoundingBoxUV bbB)
        {
            string _;
            return EstimateOverlapRatioAonB(a, b, bbA, bbB, out _);
        }

        private static double EstimateOverlapRatioAonB(
            Face a, Face b, BoundingBoxUV bbA, BoundingBoxUV bbB, out string failReason)
        {
            failReason = null;
            if (bbA == null || bbB == null) { failReason = "bb-null"; return -1; }

            // a の 4 隅を b に投影 (ProjectBCornersToA(b, a, bbA) で a の隅を b に投影)。
            // okCount       = 投影距離 ≤ tol で「b 面上にある」と判定された隅数
            // offFaceCount  = 投影は成功したが距離 > tol (b 面の外側にはみ出している)
            // exceptionCount= Evaluate/Project が例外/null (curved face 等の API 失敗)
            int okCount, offFaceCount, exceptionCount;
            BoundingBoxUV bbAonB = ProjectBCornersToA(
                b, a, bbA, out okCount, out offFaceCount, out exceptionCount);

            // 例外が 2 件以上発生 = データ不安定 → 既存ロジック (中心投影 + 面積比) を信頼
            if (exceptionCount >= 2)
            {
                failReason = "ex=" + exceptionCount;
                return -1;
            }

            // 4 隅すべてが b 面上 → AABB 交差で精密判定
            if (okCount == 4 && bbAonB != null)
            {
                double aBBuv = (bbAonB.Max.U - bbAonB.Min.U) * (bbAonB.Max.V - bbAonB.Min.V);
                if (aBBuv <= 1e-9) { failReason = "abb-zero"; return -1; }

                double iuMin = Math.Max(bbAonB.Min.U, bbB.Min.U);
                double ivMin = Math.Max(bbAonB.Min.V, bbB.Min.V);
                double iuMax = Math.Min(bbAonB.Max.U, bbB.Max.U);
                double ivMax = Math.Min(bbAonB.Max.V, bbB.Max.V);

                if (iuMax <= iuMin || ivMax <= ivMin) return 0;

                double overlapUV = (iuMax - iuMin) * (ivMax - ivMin);
                return overlapUV / aBBuv;
            }

            // 一部の隅が b の外 (= a が b より広く露出している)
            // 保守的に「on-face 隅の割合」を重なり比として返す
            //   okCount=0 → 0 (重なりほぼなし)
            //   okCount=1 → 0.25
            //   okCount=2 → 0.50
            //   okCount=3 → 0.75
            // これで閾値 0.95 を確実に下回り Stage 2 (Partial) にフォールスルーする。
            failReason = "ok=" + okCount + "/off=" + offFaceCount;
            return okCount / 4.0;
        }

        private static BoundingBoxUV ProjectBCornersToA(Face a, Face b, BoundingBoxUV bbB)
        {
            int _, __, ___;
            return ProjectBCornersToA(a, b, bbB, out _, out __, out ___);
        }

        /// <summary>
        /// 面 b の 4 隅 (BoundingBoxUV の 4 隅) を面 a に投影し、
        /// 投影先 UV を囲む AABB (Axis-Aligned Bounding Box in UV) を返す。
        /// 投影失敗や全て境界外の場合は null を返す (→ Phase 2 フォールバック)。
        ///
        /// `okCount`        = 投影距離 ≤ tol で「a 面上にある」と判定された隅数
        /// `offFaceCount`   = 投影成功したが距離 > tol (a 面から外にはみ出している)
        /// `exceptionCount` = Evaluate / Project が例外 / null (curved face 等の API 失敗)
        /// </summary>
        private static BoundingBoxUV ProjectBCornersToA(
            Face a, Face b, BoundingBoxUV bbB,
            out int okCount, out int offFaceCount, out int exceptionCount)
        {
            okCount = 0;
            offFaceCount = 0;
            exceptionCount = 0;
            var corners = new UV[]
            {
                new UV(bbB.Min.U, bbB.Min.V),
                new UV(bbB.Max.U, bbB.Min.V),
                new UV(bbB.Max.U, bbB.Max.V),
                new UV(bbB.Min.U, bbB.Max.V),
            };

            double uMin = double.MaxValue, vMin = double.MaxValue;
            double uMax = double.MinValue, vMax = double.MinValue;

            foreach (var c in corners)
            {
                XYZ p;
                try { p = b.Evaluate(c); } catch { exceptionCount++; continue; }
                if (p == null) { exceptionCount++; continue; }

                // Face.Project は「投影できない点」に対して null を返したり UVPoint=null を
                // 返すことがある (Revit API 仕様)。これは「面の外にはみ出している」という
                // 意味なので off-face として扱う (exception 扱いにすると fail-safe で
                // 誤って Full Contact 受理してしまうため)。
                IntersectionResult proj = null;
                try { proj = a.Project(p); } catch { /* swallow → off-face */ }

                if (proj == null || proj.UVPoint == null
                    || proj.Distance > CoincidenceTolFeet * 2)
                {
                    offFaceCount++;
                    continue;
                }

                var uv = proj.UVPoint;
                if (uv.U < uMin) uMin = uv.U;
                if (uv.V < vMin) vMin = uv.V;
                if (uv.U > uMax) uMax = uv.U;
                if (uv.V > vMax) vMax = uv.V;
                okCount++;
            }

            if (okCount < 3) return null; // 4隅中 3 つ以上投影できないと信頼性低い
            if (uMax <= uMin || vMax <= vMin) return null;

            return new BoundingBoxUV(uMin, vMin, uMax, vMax);
        }

        private static string Fmt(double v, int decimals)
        {
            return v.ToString("F" + decimals, CultureInfo.InvariantCulture);
        }

        private static string FmtUV(UV uv)
        {
            if (uv == null) return "-";
            return "(" + Fmt(uv.U, 3) + "," + Fmt(uv.V, 3) + ")";
        }

        private static string FmtUVBounds(BoundingBoxUV bb)
        {
            if (bb == null) return "null";
            return $"[U:{Fmt(bb.Min.U, 3)}..{Fmt(bb.Max.U, 3)},V:{Fmt(bb.Min.V, 3)}..{Fmt(bb.Max.V, 3)}]";
        }
    }
}
