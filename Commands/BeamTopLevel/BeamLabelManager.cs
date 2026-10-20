using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Tools28.Commands.BeamCommon;

namespace Tools28.Commands.BeamTopLevel
{
    /// <summary>
    /// 梁天端レベルのテキスト注釈（TextNote）を梁上に配置
    /// </summary>
    public static class BeamLabelManager
    {
        private const string LabelSuffix = "天端";

        /// <summary>
        /// 各梁の中央上方にレベル表示テキストを配置（梁方向に回転）
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

            if (overwriteExisting)
            {
                RemoveExistingBeamLabels(doc, activeView);
            }

            // 上書きしない場合は、既に同じラベルが同じ位置にある梁には作らない（再実行で重ならないように）
            var keptLabels = overwriteExisting
                ? new List<TextNote>()
                : FindExistingBeamLabels(doc, activeView);

            double textHeight = GetTextHeight(doc, textNoteTypeId);
            int viewScale = activeView.Scale;
            double modelTextHeight = textHeight * viewScale;
            double offsetDistance = modelTextHeight * 0.3;

            foreach (var beam in beams)
            {
                if (!calculationResults.ContainsKey(beam.Id))
                    continue;

                var result = calculationResults[beam.Id];
                if (!result.Success)
                    continue;

                var beamInfo = GetBeamDirectionAndMidPoint(beam, activeView);
                if (beamInfo == null)
                    continue;

                XYZ midPoint = beamInfo.Item1;
                XYZ direction = beamInfo.Item2;
                double beamWidth = beamInfo.Item3;

                XYZ normal = new XYZ(-direction.Y, direction.X, 0).Normalize();

                double totalOffset = beamWidth / 2 + offsetDistance;
                XYZ textPosition = midPoint + normal * totalOffset;

                double angle = Math.Atan2(direction.Y, direction.X);

                if (angle > Math.PI / 2 || angle < -Math.PI / 2)
                {
                    angle += Math.PI;
                    textPosition = midPoint - normal * totalOffset;
                }

                string labelText = result.DisplayValue + LabelSuffix;
                if (HasSameLabelNearby(keptLabels, labelText, textPosition, modelTextHeight))
                    continue;

                try
                {
                    TextNote textNote = TextNote.Create(doc, activeView.Id,
                        textPosition, labelText, textNoteTypeId);

                    textNote.HorizontalAlignment = HorizontalTextAlignment.Center;
                    textNote.VerticalAlignment = VerticalTextAlignment.Bottom;

                    if (Math.Abs(angle) > 0.001)
                    {
                        Line rotationAxis = Line.CreateBound(
                            textPosition,
                            textPosition + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(
                            doc, textNote.Id, rotationAxis, angle);
                    }
                }
                catch (Exception)
                {
                    // テキスト作成失敗はスキップ
                }
            }
        }

        private static Tuple<XYZ, XYZ, double> GetBeamDirectionAndMidPoint(
            FamilyInstance beam, View activeView)
        {
            var location = beam.Location as LocationCurve;
            if (location == null)
                return null;

            Curve curve = location.Curve;
            XYZ startPoint = curve.GetEndPoint(0);
            XYZ endPoint = curve.GetEndPoint(1);
            XYZ midPoint = curve.Evaluate(0.5, true);

            XYZ direction = (endPoint - startPoint);
            direction = new XYZ(direction.X, direction.Y, 0).Normalize();

            double beamWidth = BeamLabelGeometry.GetBeamWidth(beam, activeView, direction);

            midPoint = new XYZ(midPoint.X, midPoint.Y, 0);

            return Tuple.Create(midPoint, direction, beamWidth);
        }

        private static double GetTextHeight(Document doc, ElementId textNoteTypeId)
        {
            var textNoteType = doc.GetElement(textNoteTypeId) as TextNoteType;
            if (textNoteType != null)
            {
                Parameter sizeParam = textNoteType.get_Parameter(
                    BuiltInParameter.TEXT_SIZE);
                if (sizeParam != null)
                    return sizeParam.AsDouble();
            }
            return 2.5 / 304.8;
        }

        public static void RemoveExistingBeamLabels(Document doc, View activeView)
        {
            foreach (var note in FindExistingBeamLabels(doc, activeView))
            {
                try { doc.Delete(note.Id); } catch { }
            }
        }

        /// <summary>
        /// このツールが作った梁ラベル（「レベル名+数値」＋接尾辞。例: 2FL+120天端）を探す。
        /// TextNote.Text は末尾に改行が付くため取り除いてから判定する。
        /// 「スラブ天端」のような利用者の文字は、符号付きの数値が無いため対象にしない。
        /// </summary>
        private static List<TextNote> FindExistingBeamLabels(Document doc, View activeView)
        {
            var pattern = new Regex(@"[+\-]-?\d+" + Regex.Escape(LabelSuffix) + "$");
            return new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .Where(tn => pattern.IsMatch((tn.Text ?? "").TrimEnd('\r', '\n', ' ')))
                .ToList();
        }

        private static bool HasSameLabelNearby(
            List<TextNote> labels, string labelText, XYZ position, double tolerance)
        {
            foreach (var tn in labels)
            {
                if ((tn.Text ?? "").TrimEnd('\r', '\n', ' ') != labelText)
                    continue;
                XYZ c = tn.Coord;
                double dx = c.X - position.X, dy = c.Y - position.Y;
                if (Math.Sqrt(dx * dx + dy * dy) < tolerance)
                    return true;
            }
            return false;
        }
    }
}
