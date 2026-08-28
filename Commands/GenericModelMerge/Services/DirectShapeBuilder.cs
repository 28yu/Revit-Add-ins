using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Tools28.Commands.GenericModelMerge.Services
{
    /// <summary>
    /// まとめた形状を「ダイレクトシェイプ（一般モデル）」としてプロジェクトに作成する。
    ///
    /// 材質について:
    ///   Revit API には Solid の材質を直接差し替える手段がないため、作成後に
    ///   Document.Paint() で全面をペイントして単一材質の見た目にする。
    ///   面数が極端に多い場合は時間がかかるため上限を設け、超えた場合は
    ///   元の材質のままにして呼び出し側へ通知する（形状は必ず正確なまま）。
    /// </summary>
    internal static class DirectShapeBuilder
    {
        /// <summary>ペイントを行う面数の上限。これを超えたら材質適用を諦める。</summary>
        private const int MaxPaintFaces = 20000;

        internal class BuildOutcome
        {
            public ElementId ElementId { get; set; } = ElementId.InvalidElementId;
            /// <summary>材質を適用できたか。false なら元の材質のまま。</summary>
            public bool MaterialApplied { get; set; }
            /// <summary>面数が多すぎて材質適用を見送ったか。</summary>
            public bool MaterialSkippedTooManyFaces { get; set; }
        }

        /// <summary>
        /// 呼び出し側で開いているトランザクションの中で実行すること。
        /// </summary>
        public static BuildOutcome Create(
            Document doc, IList<Solid> solids, string name, ElementId materialId)
        {
            var outcome = new BuildOutcome();

            var ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            ds.ApplicationId = "Tools28";
            ds.ApplicationDataId = "GenericModelMerge";

            var geometry = new List<GeometryObject>();
            foreach (var s in solids) geometry.Add(s);
            ds.SetShape(geometry);

            if (!string.IsNullOrWhiteSpace(name))
            {
                try { ds.Name = name; } catch { }
            }

            outcome.ElementId = ds.Id;

            if (materialId != null && materialId != ElementId.InvalidElementId)
                ApplyMaterial(doc, ds, materialId, outcome);

            return outcome;
        }

        private static void ApplyMaterial(
            Document doc, DirectShape ds, ElementId materialId, BuildOutcome outcome)
        {
            // Paint するには「参照付き」の面が必要なので ComputeReferences=true で読み直す
            doc.Regenerate();

            var faces = new List<Face>();
            var opt = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = false,
                DetailLevel = ViewDetailLevel.Fine,
            };

            GeometryElement geom;
            try { geom = ds.get_Geometry(opt); }
            catch { geom = null; }
            if (geom == null) return;

            CollectFaces(geom, faces);

            if (faces.Count > MaxPaintFaces)
            {
                outcome.MaterialSkippedTooManyFaces = true;
                return;
            }

            int painted = 0;
            foreach (var f in faces)
            {
                try
                {
                    doc.Paint(ds.Id, f, materialId);
                    painted++;
                }
                catch { }
            }
            outcome.MaterialApplied = painted > 0;
        }

        private static void CollectFaces(GeometryElement geom, List<Face> output)
        {
            foreach (GeometryObject obj in geom)
            {
                if (obj is Solid s)
                {
                    foreach (Face f in s.Faces) output.Add(f);
                }
                else if (obj is GeometryInstance gi)
                {
                    var inst = gi.GetInstanceGeometry();
                    if (inst != null) CollectFaces(inst, output);
                }
            }
        }
    }
}
