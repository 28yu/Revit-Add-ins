using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

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

            double beamWidth = GetBeamWidthFromParameters(beam);

            if (beamWidth <= 0)
            {
                BoundingBoxXYZ bb = beam.get_BoundingBox(activeView);
                if (bb != null)
                {
                    XYZ normal = new XYZ(-direction.Y, direction.X, 0);

                    XYZ[] corners = new[]
                    {
                        new XYZ(bb.Min.X, bb.Min.Y, 0),
                        new XYZ(bb.Max.X, bb.Min.Y, 0),
                        new XYZ(bb.Min.X, bb.Max.Y, 0),
                        new XYZ(bb.Max.X, bb.Max.Y, 0)
                    };

                    double minProj = double.MaxValue;
                    double maxProj = double.MinValue;
                    foreach (var corner in corners)
                    {
                        double proj = corner.X * normal.X + corner.Y * normal.Y;
                        minProj = Math.Min(minProj, proj);
                        maxProj = Math.Max(maxProj, proj);
                    }

                    beamWidth = maxProj - minProj;
                }
            }

            if (beamWidth < 200.0 / 304.8)
                beamWidth = 200.0 / 304.8;

            midPoint = new XYZ(midPoint.X, midPoint.Y, 0);

            return Tuple.Create(midPoint, direction, beamWidth);
        }

        private static double GetBeamWidthFromParameters(FamilyInstance beam)
        {
            string[] widthParamNames = { "b", "B", "幅", "梁幅", "W", "w", "Width", "width" };

            foreach (string name in widthParamNames)
            {
                Parameter param = beam.LookupParameter(name);
                if (param != null && param.StorageType == StorageType.Double)
                {
                    double val = param.AsDouble();
                    if (val > 0) return val;
                }
            }

            ElementType beamType = beam.Symbol;
            if (beamType != null)
            {
                foreach (string name in widthParamNames)
                {
                    Parameter param = beamType.LookupParameter(name);
                    if (param != null && param.StorageType == StorageType.Double)
                    {
                        double val = param.AsDouble();
                        if (val > 0) return val;
                    }
                }
            }

            return 0;
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
