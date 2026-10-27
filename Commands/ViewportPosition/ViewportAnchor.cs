using System;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.ViewportPosition
{
    /// <summary>
    /// ビューポートの「図の位置」を表す情報。
    /// モデルの原点 (0,0,0) がシート上のどこに写っているか（OriginOnSheet）で位置を比べる。
    /// </summary>
    public class ViewportAnchorInfo
    {
        /// <summary>モデルの原点が写るシート上の点。計算できなかったときは null。</summary>
        public XYZ OriginOnSheet { get; set; }

        /// <summary>ビューポートの枠の中心（図の位置を計算できないときの代わり）。</summary>
        public XYZ BoxCenter { get; set; }

        public XYZ ViewDirection { get; set; }
        public XYZ UpDirection { get; set; }
        public int Scale { get; set; }

        /// <summary>パース（遠近法）の 3D ビュー。図の位置を計算できず、枠の中心でそろえても意味が無い。</summary>
        public bool IsPerspective { get; set; }

        public string ViewName { get; set; }
    }

    /// <summary>
    /// ビューポートを「枠の中心」ではなく「図の位置」でそろえるための計算。
    ///
    /// 枠の大きさはビューに写っているもの（トリミング領域が無効なら通り芯の符号・タグなども含む）で
    /// 決まるため、枠の中心をそろえても図はそろわない。そこで、モデルの原点が写るシート上の点を
    /// 「モデル → ビューの投影面 → シート」の変換で求め、その点どうしを合わせる。
    /// 平面・断面・立面・平行投影の 3D ビューで使える（見る向きと縮尺が同じであること）。
    /// </summary>
    public static class ViewportAnchor
    {
        private const double DirectionTolerance = 1e-6;

        public static ViewportAnchorInfo Capture(Document doc, Viewport vp)
        {
            var info = new ViewportAnchorInfo();
            try { info.BoxCenter = vp.GetBoxCenter(); } catch { }

            var view = doc.GetElement(vp.ViewId) as View;
            if (view == null) return info;

            info.ViewName = view.Name;
            try { info.Scale = view.Scale; } catch { }
            try { info.ViewDirection = view.ViewDirection; } catch { }
            try { info.UpDirection = view.UpDirection; } catch { }

            if (view is View3D v3 && v3.IsPerspective)
            {
                info.IsPerspective = true;
                return info;
            }

#if REVIT2021 || REVIT2022
            // Revit 2021・2022 には View.GetModelToProjectionTransforms / Viewport.GetProjectionToSheetTransform が
            // 無い（2023 で追加。Nice3point の各版でコンパイルして確認）。図の位置は計算せず、
            // 枠の中心でそろえる（呼び出し側で利用者に知らせる）
#else
            // パースビューでは例外になるので上で除外している
            try
            {
                var transforms = view.GetModelToProjectionTransforms();
                if (transforms != null && transforms.Count > 0)
                {
                    Transform modelToProjection = transforms[0].GetModelToProjectionTransform();
                    Transform projectionToSheet = vp.GetProjectionToSheetTransform();
                    XYZ p = projectionToSheet.OfPoint(modelToProjection.OfPoint(XYZ.Zero));
                    info.OriginOnSheet = new XYZ(p.X, p.Y, 0);
                }
            }
            catch (Exception ex)
            {
                // 製図ビュー・凡例など、モデルの座標を持たないビューでは計算できない
                DiagLog.Write($"[ViewportPos] 図の位置を計算できない view={view.Name}: {ex.Message}");
            }
#endif
            return info;
        }

        /// <summary>
        /// 図の位置でそろえられない組み合わせなら理由を返す（そろえられるなら null）。
        /// </summary>
        public static string Incompatibility(ViewportAnchorInfo src, ViewportAnchorInfo dst)
        {
            if (src.IsPerspective || dst.IsPerspective)
                return Loc.S("ViewportPos.Reason.Perspective");
            if (src.Scale != dst.Scale)
                return string.Format(Loc.S("ViewportPos.Reason.Scale"), src.Scale, dst.Scale);
            if (!Same(src.ViewDirection, dst.ViewDirection) || !Same(src.UpDirection, dst.UpDirection))
                return Loc.S("ViewportPos.Reason.Direction");
            return null;
        }

        private static bool Same(XYZ a, XYZ b)
        {
            if (a == null || b == null) return a == null && b == null;
            return a.IsAlmostEqualTo(b, DirectionTolerance);
        }
    }
}
