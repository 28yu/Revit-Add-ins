using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamUnderLevel
{
    /// <summary>
    /// 梁下端レベルのテキスト注釈（TextNote）を梁上に配置
    /// </summary>
    public static class BeamLabelManager
    {
        // 梁ラベルのプレフィックス（既存ラベル削除用）
        private const string LabelSuffix = "下端";

        /// <summary>
        /// 各梁の中央にレベル表示テキストを配置
        /// </summary>
        public static void CreateBeamLabels(
            Document doc,
            View activeView,
            List<FamilyInstance> beams,
            Dictionary<ElementId, BeamCalculationResult> calculationResults,
            ElementId textNoteTypeId,
            bool overwriteExisting)
        {
            if (textNoteTypeId == null)
                return;

            // 既存の梁ラベルを削除（上書きモード）
            if (overwriteExisting)
            {
                RemoveExistingBeamLabels(doc, activeView);
            }

            // 各梁にテキストを配置
            foreach (var beam in beams)
            {
                if (!calculationResults.ContainsKey(beam.Id))
                    continue;

                var result = calculationResults[beam.Id];
                if (!result.Success)
                    continue;

                // 梁の中央点を取得
                XYZ midPoint = GetBeamMidPoint(beam);
                if (midPoint == null)
                    continue;

                // テキスト内容（例: "B1FL+3300下端"）
                string labelText = result.DisplayValue + LabelSuffix;

                try
                {
                    TextNote.Create(doc, activeView.Id, midPoint,
                        labelText, textNoteTypeId);
                }
                catch (Exception)
                {
                    // テキスト作成失敗はスキップ
                }
            }
        }

        /// <summary>
        /// 梁の中央点を取得
        /// </summary>
        private static XYZ GetBeamMidPoint(FamilyInstance beam)
        {
            var location = beam.Location as LocationCurve;
            if (location != null)
            {
                Curve curve = location.Curve;
                return curve.Evaluate(0.5, true);
            }

            // LocationCurveが無い場合はBoundingBoxの中心
            BoundingBoxXYZ bb = beam.get_BoundingBox(null);
            if (bb != null)
            {
                return new XYZ(
                    (bb.Min.X + bb.Max.X) / 2,
                    (bb.Min.Y + bb.Max.Y) / 2,
                    0);
            }

            return null;
        }

        /// <summary>
        /// 既存の梁ラベルを削除
        /// </summary>
        private static void RemoveExistingBeamLabels(Document doc, View activeView)
        {
            var existingNotes = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .Where(tn => tn.Text.EndsWith(LabelSuffix))
                .ToList();

            foreach (var note in existingNotes)
            {
                try { doc.Delete(note.Id); } catch { }
            }
        }
    }
}
