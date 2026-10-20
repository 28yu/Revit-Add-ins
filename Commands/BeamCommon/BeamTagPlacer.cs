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

            foreach (var beam in targets)
            {
                // 上書きしない場合、既に同じタイプのタグがある梁には付けない（再実行で重ならないように）
                if (existing.ContainsKey(beam.Id))
                    continue;
                if (!(beam.Location is LocationCurve lc) || lc.Curve == null)
                    continue;

                Curve curve = lc.Curve;
                XYZ mid = curve.Evaluate(0.5, true);
                XYZ dir = curve.GetEndPoint(1) - curve.GetEndPoint(0);
                double angle = Math.Atan2(dir.Y, dir.X);
                // 文字が逆さにならないよう -90°〜+90° に収める
                if (angle > Math.PI / 2 + 1e-6) angle -= Math.PI;
                else if (angle < -Math.PI / 2 + 1e-6) angle += Math.PI;

                try
                {
#if REVIT2021
                    TagOrientation orientation = Math.Abs(angle) > Math.PI / 4
                        ? TagOrientation.Vertical : TagOrientation.Horizontal;
                    IndependentTag tag = IndependentTag.Create(doc, tagTypeId, view.Id,
                        new Reference(beam), false, orientation, mid);
#else
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
                    if (tag != null)
                        result.Created++;
                }
                catch (Exception)
                {
                    // タグを付けられない梁はスキップ
                }
            }

            return result;
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
