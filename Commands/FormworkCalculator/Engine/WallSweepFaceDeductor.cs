using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.FormworkCalculator.Models;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// 壁スイープ・リビール (WallSweep) の影響を受ける Wall 面を直接 DeductedContact に変更する。
    ///
    /// ContactFaceDetector の通常ロジックは「両者ともソリッドを持つ」前提だが、
    /// Reveal は壁を「切り取る」だけで自身のソリッドが無い / 不安定なケースが多い。
    /// このため Wall ↔ WallSweep 間の接触検出が漏れ、reveal で切り取られた箇所が
    /// FormworkRequired のまま残る問題があった。
    ///
    /// 本クラスでは WallSweep の BoundingBox 内に中心が含まれる Wall 面を一律
    /// DeductedContact に変更する。これは厳密ではないが、壁スイープ・リビールに
    /// 関連する付帯部の formwork 重複を防ぐのに十分。
    /// </summary>
    internal static class WallSweepFaceDeductor
    {
        private const double BoxToleranceFeet = 0.05; // 約 15mm

        internal static void DeductWallFacesNearSweeps(
            Document doc,
            List<ContactFaceDetector.ElementFacesContext> contexts,
            Dictionary<int, ElementSource> srcByContext = null)
        {
            if (contexts == null || contexts.Count == 0) return;

            // WallSweep / Wall を分離 (registry 経由 → fallback doc.GetElement)
            // 各 context にソース名 (ホスト名 or リンク名) を付与し、同じソース内でのみ
            // sweep⇔wall を組み合わせる (ホスト sweep がリンク wall を控除する誤動作を防ぐ)。
            var sweepCtxs = new List<ContactFaceDetector.ElementFacesContext>();
            var wallCtxs = new List<ContactFaceDetector.ElementFacesContext>();
            var sourceByCtxId = new Dictionary<int, string>();
            foreach (var ctx in contexts)
            {
                Element elem = null;
                string sourceName = ElementSourceRegistry.HostSourceName;
                if (srcByContext != null && srcByContext.TryGetValue(ctx.ElementId, out var src))
                {
                    elem = src.Element;
                    sourceName = src.SourceName ?? ElementSourceRegistry.HostSourceName;
                }
                if (elem == null)
                {
                    try { elem = doc.GetElement(new ElementId(ctx.ElementId)); } catch { }
                }
                sourceByCtxId[ctx.ElementId] = sourceName;
                if (elem is WallSweep) sweepCtxs.Add(ctx);
                else if (elem is Wall) wallCtxs.Add(ctx);
            }

            if (sweepCtxs.Count == 0 || wallCtxs.Count == 0)
            {
                FormworkDebugLog.Log(
                    $"  [WallSweepDeduct] sweeps={sweepCtxs.Count} walls={wallCtxs.Count} - skip");
                return;
            }

            int faceChanges = 0;
            FormworkDebugLog.Section(
                $"WallSweep Host Face Deduction (sweeps={sweepCtxs.Count} walls={wallCtxs.Count})");

            foreach (var sCtx in sweepCtxs)
            {
                if (sCtx.BB == null) continue;
                var sBb = ExpandBox(sCtx.BB, BoxToleranceFeet);
                string sSrc = sourceByCtxId.TryGetValue(sCtx.ElementId, out var sn)
                    ? sn : ElementSourceRegistry.HostSourceName;

                // リンク要素は対象外 (リンク内の wall は wall solid 自体に reveal カットが
                // 反映されている前提で、groove 面を FormworkRequired のまま残す)。
                // ホスト sweep ↔ host wall のペアのみ控除を実行。
                if (sSrc != ElementSourceRegistry.HostSourceName) continue;

                foreach (var wCtx in wallCtxs)
                {
                    if (wCtx.BB == null) continue;
                    // ソースが異なる場合はスキップ (host sweep ↔ link wall の誤マッチを防ぐ)
                    string wSrc = sourceByCtxId.TryGetValue(wCtx.ElementId, out var wn)
                        ? wn : ElementSourceRegistry.HostSourceName;
                    if (sSrc != wSrc) continue;
                    if (!BoxesOverlap(sBb, wCtx.BB)) continue;

                    int wallChanges = 0;
                    foreach (var fi in wCtx.Faces)
                    {
                        if (fi.FaceType != FaceType.FormworkRequired) continue;
                        if (fi.Face == null) continue;

                        // 面の中心点を取得
                        XYZ center = null;
                        try
                        {
                            var bbUv = fi.Face.GetBoundingBox();
                            if (bbUv != null)
                                center = fi.Face.Evaluate((bbUv.Min + bbUv.Max) * 0.5);
                        }
                        catch { }
                        if (center == null) continue;

                        if (IsPointInBox(center, sBb))
                        {
                            fi.FaceType = FaceType.DeductedContact;
                            faceChanges++;
                            wallChanges++;
                        }
                    }
                    if (wallChanges > 0)
                    {
                        FormworkDebugLog.Log(
                            $"  [WallSweepDeduct] sweep E{sCtx.ElementId} → wall E{wCtx.ElementId}: " +
                            $"{wallChanges} faces deducted");
                    }
                }
            }

            FormworkDebugLog.Log($"  [WallSweepDeduct] total face changes: {faceChanges}");
            FormworkDebugLog.Flush();
        }

        private static BoundingBoxXYZ ExpandBox(BoundingBoxXYZ b, double tol)
        {
            return new BoundingBoxXYZ
            {
                Min = new XYZ(b.Min.X - tol, b.Min.Y - tol, b.Min.Z - tol),
                Max = new XYZ(b.Max.X + tol, b.Max.Y + tol, b.Max.Z + tol),
            };
        }

        private static bool BoxesOverlap(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            return a.Min.X <= b.Max.X && b.Min.X <= a.Max.X
                && a.Min.Y <= b.Max.Y && b.Min.Y <= a.Max.Y
                && a.Min.Z <= b.Max.Z && b.Min.Z <= a.Max.Z;
        }

        private static bool IsPointInBox(XYZ p, BoundingBoxXYZ b)
        {
            if (p == null || b == null) return false;
            return p.X >= b.Min.X && p.X <= b.Max.X
                && p.Y >= b.Min.Y && p.Y <= b.Max.Y
                && p.Z >= b.Min.Z && p.Z <= b.Max.Z;
        }
    }
}
