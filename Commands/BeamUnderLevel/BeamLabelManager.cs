using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamUnderLevel
{
    /// <summary>
    /// 梁下端レベルのテキスト注釈（TextNote）を梁上に配置
    /// </summary>
    public static class BeamLabelManager
    {
        private const string LabelSuffix = "下端";

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

            // 既存の梁ラベルを削除（上書きモード）
            if (overwriteExisting)
            {
                RemoveExistingBeamLabels(doc, activeView);
            }

            // 上書きしない場合は、既に同じラベルが同じ位置にある梁には作らない（再実行で重ならないように）
            var keptLabels = overwriteExisting
                ? new List<TextNote>()
                : FindExistingBeamLabels(doc, activeView);

            // テキストサイズを取得（梁の上にオフセットするため）
            double textHeight = GetTextHeight(doc, textNoteTypeId);
            // ビューのスケールを考慮してモデル空間でのテキスト高さを計算
            // TEXT_SIZE はペーパー空間の高さなので、モデル空間では viewScale 倍
            int viewScale = activeView.Scale;
            double modelTextHeight = textHeight * viewScale;
            // オフセット量: 梁端からの余白（テキストは下端基準なので上に伸びる）
            double offsetDistance = modelTextHeight * 0.3;

            // 各梁にテキストを配置
            foreach (var beam in beams)
            {
                if (!calculationResults.ContainsKey(beam.Id))
                    continue;

                var result = calculationResults[beam.Id];
                if (!result.Success)
                    continue;

                // 梁の方向と中央点を取得
                var beamInfo = GetBeamDirectionAndMidPoint(beam, activeView);
                if (beamInfo == null)
                    continue;

                XYZ midPoint = beamInfo.Item1;
                XYZ direction = beamInfo.Item2;
                double beamWidth = beamInfo.Item3;

                // 梁の法線方向（上方向）にオフセット
                // 天井伏図はXY平面なので、梁方向を90度回転させた方向が法線
                XYZ normal = new XYZ(-direction.Y, direction.X, 0).Normalize();

                // テキスト位置: 梁中央から法線方向にオフセット
                double totalOffset = beamWidth / 2 + offsetDistance;
                XYZ textPosition = midPoint + normal * totalOffset;

                // 梁方向の角度を計算（ラジアン）
                double angle = Math.Atan2(direction.Y, direction.X);

                // テキストが上下逆にならないよう調整
                // 角度が90°〜270°（左向き）の場合は180°回転
                if (angle > Math.PI / 2 || angle < -Math.PI / 2)
                {
                    angle += Math.PI;
                    // オフセット方向も反転
                    textPosition = midPoint - normal * totalOffset;
                }

                string labelText = result.DisplayValue + LabelSuffix;
                if (HasSameLabelNearby(keptLabels, labelText, textPosition, modelTextHeight))
                    continue;

                try
                {
                    TextNote textNote = TextNote.Create(doc, activeView.Id,
                        textPosition, labelText, textNoteTypeId);

                    // 位置合わせ: 中心・下（アンカーがテキスト下端中央になり、梁から離れる方向にテキストが伸びる）
                    textNote.HorizontalAlignment = HorizontalTextAlignment.Center;
                    textNote.VerticalAlignment = VerticalTextAlignment.Bottom;

                    // 梁方向に回転
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

        /// <summary>
        /// 梁の方向ベクトル、中央点、梁幅を取得
        /// Returns: (midPoint, direction, beamWidth)
        /// </summary>
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

            // 梁方向ベクトル（XY平面上）
            XYZ direction = (endPoint - startPoint);
            direction = new XYZ(direction.X, direction.Y, 0).Normalize();

            // 梁幅を取得（タイプ/インスタンスパラメータから直接取得を優先）
            double beamWidth = GetBeamWidthFromParameters(beam);

            // パラメータから取得できない場合はBoundingBoxからフォールバック
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

            // 最小幅を保証（200mm = 約0.656ft）
            if (beamWidth < 200.0 / 304.8)
                beamWidth = 200.0 / 304.8;

            // Z=0 にする（天井伏図はXY平面）
            midPoint = new XYZ(midPoint.X, midPoint.Y, 0);

            return Tuple.Create(midPoint, direction, beamWidth);
        }

        /// <summary>
        /// 梁のタイプ/インスタンスパラメータから梁幅を取得
        /// </summary>
        private static double GetBeamWidthFromParameters(FamilyInstance beam)
        {
            // よく使われる梁幅パラメータ名
            string[] widthParamNames = { "b", "B", "幅", "梁幅", "W", "w", "Width", "width" };

            // インスタンスパラメータを検索
            foreach (string name in widthParamNames)
            {
                Parameter param = beam.LookupParameter(name);
                if (param != null && param.StorageType == StorageType.Double)
                {
                    double val = param.AsDouble();
                    if (val > 0) return val;
                }
            }

            // タイプパラメータを検索
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

        /// <summary>
        /// TextNoteTypeからテキスト高さを取得
        /// </summary>
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

        /// <summary>
        /// 既存の梁ラベルを削除
        /// </summary>
        private static void RemoveExistingBeamLabels(Document doc, View activeView)
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
