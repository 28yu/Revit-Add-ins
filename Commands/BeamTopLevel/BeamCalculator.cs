using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamTopLevel
{
    /// <summary>
    /// 梁天端レベル計算結果
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
    /// 梁天端レベル取得ロジック
    /// </summary>
    public static class BeamCalculator
    {
        // 梁天端レベルパラメータの検索候補（優先度順）
        private static readonly string[] TopLevelParamNames = new[]
        {
            "始端レベル オフセット", "終端レベル オフセット",
            "始端レベルオフセット", "終端レベルオフセット",
            "Start Level Offset", "End Level Offset",
            "天端レベル", "Top Level", "Top Offset",
            "レベルからのオフセット値", "Offset from Level"
        };

        /// <summary>
        /// ファミリ毎に梁天端レベルパラメータ候補を検索
        /// </summary>
        public static Dictionary<string, List<ParamCandidate>> FindTopLevelParameterCandidates(
            Dictionary<string, List<FamilyInstance>> beamsByFamily)
        {
            var result = new Dictionary<string, List<ParamCandidate>>();

            foreach (var entry in beamsByFamily)
            {
                string familyName = entry.Key;
                var beamList = entry.Value;
                var candidates = new List<ParamCandidate>();

                // BuiltInParameter で検索（始端/終端レベルオフセット）
                var builtInParams = new[]
                {
                    new { Param = BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION, Name = "始端レベル オフセット" },
                    new { Param = BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION, Name = "終端レベル オフセット" }
                };

                foreach (var bip in builtInParams)
                {
                    int detectedCount = 0;
                    foreach (var beam in beamList)
                    {
                        Parameter param = beam.get_Parameter(bip.Param);
                        if (param != null && param.HasValue)
                            detectedCount++;
                    }

                    if (detectedCount > 0)
                    {
                        candidates.Add(new ParamCandidate
                        {
                            ParamName = bip.Name,
                            DetectedCount = detectedCount
                        });
                    }
                }

                // 名前で検索
                foreach (string paramName in TopLevelParamNames)
                {
                    if (candidates.Any(c => c.ParamName == paramName))
                        continue;

                    int detectedCount = 0;
                    foreach (var beam in beamList)
                    {
                        Parameter param = beam.LookupParameter(paramName);
                        if (param != null && param.HasValue &&
                            param.StorageType == StorageType.Double)
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

                result[familyName] = candidates;
            }

            return result;
        }

        /// <summary>
        /// ファミリ毎にDouble型パラメータを収集（レベル・オフセット関連に絞る）
        /// 主要候補以外の追加パラメータ用
        /// </summary>
        public static Dictionary<string, List<string>> FindAdditionalLevelParameters(
            Dictionary<string, List<FamilyInstance>> beamsByFamily,
            Dictionary<string, List<ParamCandidate>> alreadyFound)
        {
            string[] levelKeywords = new[]
            {
                "レベル", "オフセット", "天端", "上端", "下端",
                "level", "offset", "elevation", "top", "bottom"
            };

            var result = new Dictionary<string, List<string>>();

            foreach (var entry in beamsByFamily)
            {
                string familyName = entry.Key;
                var beamList = entry.Value;
                var existingNames = alreadyFound.ContainsKey(familyName)
                    ? alreadyFound[familyName].Select(c => c.ParamName).ToHashSet()
                    : new HashSet<string>();

                var additionalParams = new HashSet<string>();

                if (beamList.Count > 0)
                {
                    var sampleBeam = beamList[0];

                    // インスタンスパラメータをスキャン
                    foreach (Parameter param in sampleBeam.Parameters)
                    {
                        if (param.StorageType != StorageType.Double || !param.HasValue)
                            continue;

                        string name = param.Definition.Name;
                        if (existingNames.Contains(name))
                            continue;

                        string nameLower = name.ToLower();
                        if (levelKeywords.Any(kw => nameLower.Contains(kw.ToLower())))
                        {
                            additionalParams.Add(name);
                        }
                    }

                    // タイプパラメータもスキャン
                    FamilySymbol symbol = sampleBeam.Symbol;
                    if (symbol != null)
                    {
                        foreach (Parameter param in symbol.Parameters)
                        {
                            if (param.StorageType != StorageType.Double || !param.HasValue)
                                continue;

                            string name = param.Definition.Name;
                            if (existingNames.Contains(name) || additionalParams.Contains(name))
                                continue;

                            string nameLower = name.ToLower();
                            if (levelKeywords.Any(kw => nameLower.Contains(kw.ToLower())))
                            {
                                additionalParams.Add(name);
                            }
                        }
                    }
                }

                result[familyName] = additionalParams.OrderBy(n => n).ToList();
            }

            return result;
        }

        /// <summary>
        /// 梁天端レベルを取得（参照レベル基準のパラメータ値をそのまま使用）
        /// </summary>
        public static BeamCalculationResult Calculate(
            FamilyInstance beam,
            string refLevelName,
            string topLevelParamName)
        {
            try
            {
                // 天端レベルパラメータの値を取得（参照レベル基準、内部単位 feet）
                double topLevel = GetBeamTopLevel(beam, topLevelParamName);

                // mm単位に変換して表示形式を生成
                double topLevelMm = Math.Round(FeetToMm(topLevel));

                string sign = topLevelMm >= 0 ? "+" : "";
                string displayValue = $"{refLevelName}{sign}{topLevelMm:0}";

                return new BeamCalculationResult
                {
                    Success = true,
                    RefLevelName = refLevelName,
                    LevelDifference = topLevel,
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
        /// 梁天端レベル（オフセット）を取得
        /// </summary>
        private static double GetBeamTopLevel(FamilyInstance beam, string paramName)
        {
            // BuiltInParameter で取得を試行（始端/終端レベルオフセット）
            if (paramName == "始端レベル オフセット" || paramName == "始端レベルオフセット" ||
                paramName == "Start Level Offset")
            {
                Parameter bipParam = beam.get_Parameter(
                    BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION);
                if (bipParam != null && bipParam.HasValue)
                    return bipParam.AsDouble();
            }
            else if (paramName == "終端レベル オフセット" || paramName == "終端レベルオフセット" ||
                     paramName == "End Level Offset")
            {
                Parameter bipParam = beam.get_Parameter(
                    BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION);
                if (bipParam != null && bipParam.HasValue)
                    return bipParam.AsDouble();
            }

            // 名前で検索
            Parameter param = beam.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                return param.AsDouble();

            return 0;
        }

        /// <summary>
        /// フィートからmm単位への変換ヘルパー
        /// </summary>
        public static double FeetToMm(double feet)
        {
            return UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Millimeters);
        }

        /// <summary>
        /// mmからフィートへの変換ヘルパー
        /// </summary>
        public static double MmToFeet(double mm)
        {
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
        }
    }
}
