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
        /// ファミリ毎に梁高さ用の追加パラメータを収集（Double型、値>0のもの）
        /// 主要候補以外の追加パラメータ用
        /// </summary>
        public static Dictionary<string, List<string>> FindAdditionalHeightParameters(
            Dictionary<string, List<FamilyInstance>> beamsByFamily,
            Dictionary<string, List<ParamCandidate>> alreadyFound)
        {
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
                        if (param.AsDouble() <= 0)
                            continue;

                        string name = param.Definition.Name;
                        if (existingNames.Contains(name))
                            continue;

                        additionalParams.Add(name);
                    }

                    // タイプパラメータもスキャン
                    FamilySymbol symbol = sampleBeam.Symbol;
                    if (symbol != null)
                    {
                        foreach (Parameter param in symbol.Parameters)
                        {
                            if (param.StorageType != StorageType.Double || !param.HasValue)
                                continue;
                            if (param.AsDouble() <= 0)
                                continue;

                            string name = param.Definition.Name;
                            if (existingNames.Contains(name) || additionalParams.Contains(name))
                                continue;

                            additionalParams.Add(name);
                        }
                    }
                }

                result[familyName] = additionalParams.OrderBy(n => n).ToList();
            }

            return result;
        }

        /// <summary>
        /// ファミリ毎に梁のDouble型パラメータを収集（レベル・オフセット関連に絞る）
        /// 主要候補以外の追加パラメータ用
        /// </summary>
        public static Dictionary<string, List<string>> FindAdditionalLevelParameters(
            Dictionary<string, List<FamilyInstance>> beamsByFamily,
            Dictionary<string, List<ParamCandidate>> alreadyFound)
        {
            // レベル・オフセット関連のキーワード
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
        /// 梁下端レベルを計算
        /// 計算式: 梁下端レベル = 階高 - 梁天端レベル - 梁高さ
        ///   （梁天端レベルは上位レベルからの下がり量）
        ///   コード上: bottomLevel = floorHeight + topLevelOffset - beamHeight
        ///   （topLevelOffsetは負値＝上位レベルからの下がり）
        /// </summary>
        public static BeamCalculationResult Calculate(
            FamilyInstance beam,
            double floorHeight,
            string refLevelName,
            string heightParamName,
            string topLevelParamName)
        {
            try
            {
                // 1. 梁天端レベル（上位レベルからのオフセット、通常は負値）を取得
                double topLevelOffset = GetBeamTopLevel(beam, topLevelParamName);

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

                // 3. 梁下端レベルを計算（フィート単位、参照レベル基準）
                //    梁天端 = 階高 + オフセット（例: 3000 + (-300) = 2700）
                //    梁下端 = 梁天端 - 梁高さ
                double bottomLevel = floorHeight + topLevelOffset - beamHeight;

                // 4. mm単位に変換して表示形式を生成
                double bottomLevelMm = Math.Round(FeetToMm(bottomLevel));

                string sign = bottomLevelMm >= 0 ? "+" : "";
                string displayValue = $"{refLevelName}{sign}{bottomLevelMm:0}";

                return new BeamCalculationResult
                {
                    Success = true,
                    RefLevelName = refLevelName,
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
            // UnitTypeId.Millimeters は Revit 2021+ で使用可能
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
