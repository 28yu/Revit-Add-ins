using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamUnderLevel
{
    /// <summary>
    /// 梁下端レベル計算結果
    /// </summary>
    public class BeamCalculationResult
    {
        public bool Success { get; set; }
        public string RefLevelName { get; set; }
        public double LevelDifference { get; set; }
        public string DisplayValue { get; set; }
        public string Error { get; set; }
    }

    /// <summary>
    /// ファミリ毎のパラメータ候補情報
    /// </summary>
    public class ParamCandidate
    {
        public string ParamName { get; set; }
        public int DetectedCount { get; set; }
    }

    /// <summary>
    /// 梁下端レベル計算ロジック
    /// </summary>
    public static class BeamCalculator
    {
        // 梁高さパラメータの検索候補（優先度順）
        private static readonly string[] HeightParamNames = new[]
        {
            "h", "H", "梁成", "高さ", "Depth", "d", "D",
            "b", "厚さ", "width", "幅", "Height"
        };

        /// <summary>
        /// ファミリ毎に梁高さパラメータ候補を検索
        /// </summary>
        public static Dictionary<string, List<ParamCandidate>> FindHeightParameterCandidates(
            Dictionary<string, List<FamilyInstance>> beamsByFamily)
        {
            var result = new Dictionary<string, List<ParamCandidate>>();

            foreach (var entry in beamsByFamily)
            {
                string familyName = entry.Key;
                var beamList = entry.Value;
                var candidates = new List<ParamCandidate>();

                foreach (string paramName in HeightParamNames)
                {
                    int detectedCount = 0;
                    foreach (var beam in beamList)
                    {
                        Parameter param = beam.LookupParameter(paramName);
                        if (param != null && param.HasValue &&
                            param.StorageType == StorageType.Double &&
                            param.AsDouble() > 0)
                        {
                            detectedCount++;
                        }
                    }

                    if (detectedCount > 0)
                    {
                        candidates.Add(new ParamCandidate
                        {
                            ParamName = paramName,
                            DetectedCount = detectedCount
                        });
                    }
                }

                // タイプパラメータも検索（FamilySymbol経由）
                if (beamList.Count > 0)
                {
                    FamilySymbol symbol = beamList[0].Symbol;
                    foreach (string paramName in HeightParamNames)
                    {
                        // 既にインスタンスパラメータで見つかっている場合はスキップ
                        if (candidates.Any(c => c.ParamName == paramName))
                            continue;

                        Parameter typeParam = symbol.LookupParameter(paramName);
                        if (typeParam != null && typeParam.HasValue &&
                            typeParam.StorageType == StorageType.Double &&
                            typeParam.AsDouble() > 0)
                        {
                            candidates.Add(new ParamCandidate
                            {
                                ParamName = paramName,
                                DetectedCount = beamList.Count
                            });
                        }
                    }
                }

                result[familyName] = candidates;
            }

            return result;
        }

        /// <summary>
        /// 梁下端レベルを計算
        /// 計算式: 梁下端レベル = 階高 - 始端レベルオフセット - 梁高さ
        /// </summary>
        public static BeamCalculationResult Calculate(
            FamilyInstance beam,
            double floorHeight,
            string lowerLevelName,
            string heightParamName)
        {
            try
            {
                // 1. 始端レベルオフセットを取得
                double offset = GetStartLevelOffset(beam);

                // 2. 梁高さを取得
                double beamHeight = GetBeamHeight(beam, heightParamName);
                if (beamHeight <= 0)
                {
                    return new BeamCalculationResult
                    {
                        Success = false,
                        Error = $"梁高さパラメータ「{heightParamName}」の値が無効 (≤0)"
                    };
                }

                // 3. 梁下端レベルを計算（フィート単位）
                double bottomLevel = floorHeight - offset - beamHeight;

                // 4. mm単位に変換して表示形式を生成
                double bottomLevelMm = Math.Round(FeetToMm(bottomLevel));

                string sign = bottomLevelMm >= 0 ? "+" : "";
                string displayValue = $"{lowerLevelName}{sign}{bottomLevelMm:0}";

                return new BeamCalculationResult
                {
                    Success = true,
                    RefLevelName = lowerLevelName,
                    LevelDifference = bottomLevel,
                    DisplayValue = displayValue
                };
            }
            catch (Exception ex)
            {
                return new BeamCalculationResult
                {
                    Success = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 始端レベルオフセットを取得
        /// </summary>
        private static double GetStartLevelOffset(FamilyInstance beam)
        {
            // BuiltInParameter で取得を試行
            Parameter offsetParam = beam.get_Parameter(
                BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION);
            if (offsetParam != null && offsetParam.HasValue)
                return offsetParam.AsDouble();

            // 名前で検索（日本語/英語）
            string[] offsetParamNames = new[]
            {
                "始端レベル オフセット",
                "Start Level Offset",
                "始端レベルオフセット"
            };

            foreach (string name in offsetParamNames)
            {
                Parameter param = beam.LookupParameter(name);
                if (param != null && param.HasValue)
                    return param.AsDouble();
            }

            return 0;
        }

        /// <summary>
        /// 梁高さを取得（インスタンス → タイプの順で検索）
        /// </summary>
        private static double GetBeamHeight(FamilyInstance beam, string paramName)
        {
            // インスタンスパラメータから検索
            Parameter param = beam.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.Double)
            {
                double value = param.AsDouble();
                if (value > 0) return value;
            }

            // タイプパラメータから検索
            FamilySymbol symbol = beam.Symbol;
            if (symbol != null)
            {
                Parameter typeParam = symbol.LookupParameter(paramName);
                if (typeParam != null && typeParam.HasValue &&
                    typeParam.StorageType == StorageType.Double)
                {
                    double value = typeParam.AsDouble();
                    if (value > 0) return value;
                }
            }

            return 0;
        }

        /// <summary>
        /// フィートからmm単位への変換ヘルパー
        /// </summary>
        public static double FeetToMm(double feet)
        {
#if REVIT2022 || REVIT2023
            return UnitUtils.ConvertFromInternalUnits(feet, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Millimeters);
#endif
        }

        /// <summary>
        /// mmからフィートへの変換ヘルパー
        /// </summary>
        public static double MmToFeet(double mm)
        {
#if REVIT2022 || REVIT2023
            return UnitUtils.ConvertToInternalUnits(mm, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
#endif
        }
    }
}
