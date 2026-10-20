using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamCommon
{
    /// <summary>
    /// 梁タグ（構造フレームタグ）のタイプ表示用ラッパー
    /// </summary>
    public class BeamTagTypeItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }

        public BeamTagTypeItem(FamilySymbol symbol)
        {
            Id = symbol.Id;
            Name = $"{symbol.FamilyName} : {symbol.Name}";
        }

        public override string ToString()
        {
            return Name;
        }
    }

    /// <summary>
    /// 梁タグの配置結果
    /// </summary>
    public class BeamTagResult
    {
        public int Created { get; set; }
    }

    /// <summary>
    /// 梁天端・梁下端色分けで、レベルの文字の代わりに梁タグを配置する。
    /// タグに表示する内容はタグファミリ側で決まるため、梁レベルが表示されるよう
    /// 設定済みのタグを利用者が選ぶ前提とする（このツールはタグを置くだけ）。
    /// </summary>
    public static class BeamTagPlacer
    {
        /// <summary>
        /// モデルに読み込まれている構造フレームタグのタイプ一覧
        /// </summary>
        public static List<BeamTagTypeItem> CollectTagTypes(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_StructuralFramingTags)
                .Cast<FamilySymbol>()
                .Select(s => new BeamTagTypeItem(s))
                .OrderBy(t => t.Name)
                .ToList();
        }

        /// <summary>
        /// 梁タグを配置する（トランザクション内で、梁にパラメータ値を書き込んだ後に呼ぶ）
        /// </summary>
        /// <param name="targetIds">タグを付ける梁（レベルを算出できた梁）</param>
        public static BeamTagResult CreateTags(
            Document doc,
            View view,
            IEnumerable<FamilyInstance> beams,
            ICollection<ElementId> targetIds,
            ElementId tagTypeId,
            bool overwriteExisting)
        {
            var result = new BeamTagResult();
            if (tagTypeId == null || tagTypeId == ElementId.InvalidElementId)
                return result;

            var targets = beams.Where(b => targetIds.Contains(b.Id)).ToList();

            // このビューにある同じタイプのタグ（梁ID → タグ）
            var existing = new Dictionary<ElementId, List<IndependentTag>>();
            foreach (var tag in new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(IndependentTag))
                .Cast<IndependentTag>()
                .Where(t => t.GetTypeId() == tagTypeId))
            {
                foreach (ElementId id in GetTaggedIds(tag))
                {
                    if (!targetIds.Contains(id))
                        continue;
                    if (!existing.TryGetValue(id, out var list))
                        existing[id] = list = new List<IndependentTag>();
                    list.Add(tag);
                }
            }

            // 上書き: 同じタイプのタグを付け直す（他のタイプのタグ＝符号タグ等には触れない）
            if (overwriteExisting)
            {
                foreach (var tag in existing.Values.SelectMany(l => l).Distinct())
                {
                    try { doc.Delete(tag.Id); } catch (Exception) { }
                }
                existing.Clear();
            }

            // ビューの右・上方向（ビューを回転していても画面に対して読める向きにするため）
            XYZ right = view.RightDirection;
            XYZ up = view.UpDirection;

            var placed = new List<PlacedTag>();
            foreach (var beam in targets)
            {
                // 上書きしない場合、既に同じタイプのタグがある梁には付けない（再実行で重ならないように）
                if (existing.ContainsKey(beam.Id))
                    continue;
                if (!(beam.Location is LocationCurve lc) || lc.Curve == null)
                    continue;

                Curve curve = lc.Curve;
                XYZ mid = curve.Evaluate(0.5, true);
                XYZ d = curve.GetEndPoint(1) - curve.GetEndPoint(0);
                XYZ dirXY = new XYZ(d.X, d.Y, 0);
                if (dirXY.GetLength() < 1e-9)
                    continue;
                dirXY = dirXY.Normalize();

                // 画面上での梁の角度。文字が逆さにならないよう -90°〜+90° に収める
                double angle = Math.Atan2(dirXY.DotProduct(up), dirXY.DotProduct(right));
                if (angle > Math.PI / 2 + 1e-6) angle -= Math.PI;
                else if (angle < -Math.PI / 2 + 1e-6) angle += Math.PI;

                try
                {
#if REVIT2021
                    // 2021 は任意の角度に回せないため、水平／垂直のうち近い方にする
                    bool vertical = Math.Abs(angle) > Math.PI / 4;
                    double tagAngle = vertical ? Math.PI / 2 : 0;
                    IndependentTag tag = IndependentTag.Create(doc, tagTypeId, view.Id,
                        new Reference(beam), false,
                        vertical ? TagOrientation.Vertical : TagOrientation.Horizontal, mid);
#else
                    double tagAngle = angle;
                    IndependentTag tag = IndependentTag.Create(doc, tagTypeId, view.Id,
                        new Reference(beam), false, TagOrientation.Horizontal, mid);
                    if (Math.Abs(angle) > 0.001)
                    {
                        // 梁の向きに合わせて回転する（タグファミリが「構成要素と一緒に回転」でも問題ない）
                        try
                        {
                            tag.TagOrientation = TagOrientation.AnyModelDirection;
                            tag.RotationAngle = angle;
                        }
                        catch (Exception) { }
                    }
#endif
                    if (tag == null)
                        continue;
                    result.Created++;

                    // タグの読む向き（along）と、それに直交する上向き（side）。文字の梁ラベルと同じく side 側に置く
                    XYZ along = right * Math.Cos(tagAngle) + up * Math.Sin(tagAngle);
                    XYZ side = right * -Math.Sin(tagAngle) + up * Math.Cos(tagAngle);
                    placed.Add(new PlacedTag
                    {
                        Tag = tag,
                        Mid = mid,
                        Along = along,
                        Side = side,
                        HalfWidth = BeamLabelGeometry.GetBeamWidth(beam, view, dirXY) / 2
                    });
                }
                catch (Exception)
                {
                    // タグを付けられない梁はスキップ
                }
            }

            if (placed.Count > 0)
                MoveBesideBeams(doc, view, placed);

            return result;
        }

        private class PlacedTag
        {
            public IndependentTag Tag;
            public XYZ Mid;
            public XYZ Along;
            public XYZ Side;
            public double HalfWidth;
        }

        /// <summary>
        /// タグを、文字の梁ラベルと同じ位置（梁の脇で、梁の中央）に動かす。
        /// タグの大きさや基準点はタグファミリごとに違うため、置いたタグの外形を測ってから動かす。
        /// 梁と同じ向きの外形で測るので、斜めの梁では少しだけ梁から離れる
        /// </summary>
        private static void MoveBesideBeams(Document doc, View view, List<PlacedTag> placed)
        {
            try { doc.Regenerate(); } catch (Exception) { return; }

            foreach (var p in placed)
            {
                try
                {
                    BoundingBoxXYZ bb = p.Tag.get_BoundingBox(view);
                    if (bb == null)
                        continue;

                    double minA = double.MaxValue, maxA = double.MinValue;
                    double minS = double.MaxValue, maxS = double.MinValue;
                    foreach (double x in new[] { bb.Min.X, bb.Max.X })
                    foreach (double y in new[] { bb.Min.Y, bb.Max.Y })
                    {
                        XYZ v = new XYZ(x - p.Mid.X, y - p.Mid.Y, 0);
                        double a = v.DotProduct(p.Along), sd = v.DotProduct(p.Side);
                        minA = Math.Min(minA, a); maxA = Math.Max(maxA, a);
                        minS = Math.Min(minS, sd); maxS = Math.Max(maxS, sd);
                    }

                    // 余白はタグの高さの 0.3 倍（文字の梁ラベルの余白＝文字高さの 0.3 倍に合わせる）
                    double gap = (maxS - minS) * 0.3;
                    double shiftAlong = -(minA + maxA) / 2;
                    double shiftSide = p.HalfWidth + gap - minS;
                    XYZ shift = p.Along * shiftAlong + p.Side * shiftSide;
                    if (shift.GetLength() > 1e-6)
                        p.Tag.TagHeadPosition = p.Tag.TagHeadPosition + shift;
                }
                catch (Exception)
                {
                    // 動かせないタグは梁の中央に残す
                }
            }
        }

        private static IEnumerable<ElementId> GetTaggedIds(IndependentTag tag)
        {
#if REVIT2021
            ElementId id = tag.TaggedLocalElementId;
            return id != null && id != ElementId.InvalidElementId
                ? new[] { id } : Enumerable.Empty<ElementId>();
#else
            return tag.GetTaggedLocalElementIds();
#endif
        }
    }
}
