using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.Room3DColor
{
    /// <summary>
    /// 部屋3D色分けのグルーピング・色パレット・グラフィックオーバーライドを管理
    /// </summary>
    public static class RoomColorManager
    {
        /// <summary>値が未設定の場合のラベル</summary>
        public const string NoValueLabel = "(未設定)";

        /// <summary>輪郭線（投影線）用の薄いグレー</summary>
        public static readonly Color OutlineColor = new Color(190, 190, 190);

        /// <summary>
        /// 色分け基準に応じて部屋をグループ化し、各グループに色を割り当てる
        /// </summary>
        public static List<RoomColorGroup> BuildGroups(
            IEnumerable<RoomEntry> entries,
            RoomColorBasis basis,
            string parameterName)
        {
            var groups = new List<RoomColorGroup>();

            if (basis == RoomColorBasis.All)
            {
                // 全部屋を1グループにまとめる
                groups.Add(new RoomColorGroup
                {
                    Key = "All",
                    Label = Loc.S("Room3D.AllGroupLabel"),
                    RoomIds = entries.Select(e => e.Id).ToList()
                });
            }
            else if (basis == RoomColorBasis.PerRoom)
            {
                // 部屋ごとに1グループ
                foreach (var e in entries.OrderBy(e => e.Name))
                {
                    groups.Add(new RoomColorGroup
                    {
                        Key = e.Id.IntValue().ToString(),
                        Label = e.Name,
                        RoomIds = new List<ElementId> { e.Id }
                    });
                }
            }
            else
            {
                // キーごとにまとめる
                var dict = new Dictionary<string, RoomColorGroup>();
                foreach (var e in entries)
                {
                    string key = GetGroupKey(e, basis, parameterName);
                    if (!dict.TryGetValue(key, out var g))
                    {
                        g = new RoomColorGroup { Key = key, Label = key };
                        dict[key] = g;
                    }
                    g.RoomIds.Add(e.Id);
                }

                groups = dict.Values
                    .OrderBy(g => g.Label, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            // 色割当
            List<Color> colors = GenerateColors(groups.Count);
            for (int i = 0; i < groups.Count; i++)
                groups[i].Color = colors[i];

            return groups;
        }

        /// <summary>
        /// 部屋1件のグループキーを算出（BuildGroups と同一ロジック）
        /// </summary>
        public static string GetGroupKey(RoomEntry e, RoomColorBasis basis, string parameterName)
        {
            switch (basis)
            {
                case RoomColorBasis.ByName:
                    return string.IsNullOrEmpty(e.Name) ? NoValueLabel : e.Name;
                case RoomColorBasis.ByLevel:
                    return string.IsNullOrEmpty(e.LevelName) ? NoValueLabel : e.LevelName;
                case RoomColorBasis.ByParameter:
                    if (!string.IsNullOrEmpty(parameterName)
                        && e.Parameters.TryGetValue(parameterName, out var v)
                        && !string.IsNullOrEmpty(v))
                        return v;
                    return NoValueLabel;
                case RoomColorBasis.PerRoom:
                    return e.Id.IntValue().ToString();
                case RoomColorBasis.All:
                    return "All";
                default:
                    return NoValueLabel;
            }
        }

        /// <summary>
        /// ソリッド塗りつぶしパターンのIdを取得
        /// </summary>
        public static ElementId GetSolidFillPatternId(Document doc)
        {
            var fillPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            return fillPattern?.Id;
        }

        /// <summary>
        /// 3Dビュー用の色オーバーライドを作成（サーフェス前景をソリッド塗り＋色）
        /// </summary>
        public static OverrideGraphicSettings CreateColorOverrides(
            Color color, ElementId solidFillPatternId)
        {
            var overrides = new OverrideGraphicSettings();

            if (solidFillPatternId != null && solidFillPatternId != ElementId.InvalidElementId)
            {
                overrides.SetSurfaceForegroundPatternVisible(true);
                overrides.SetSurfaceForegroundPatternId(solidFillPatternId);
                overrides.SetSurfaceForegroundPatternColor(color);
            }

            // 輪郭線（投影線）は薄いグレー・最細にして目立たせない
            overrides.SetProjectionLineColor(OutlineColor);
            overrides.SetProjectionLineWeight(1);

            return overrides;
        }

        /// <summary>
        /// パステル系ベースカラーからN色を生成（BeamTopLevel と同系統のパレット）
        /// </summary>
        public static List<Color> GenerateColors(int colorCount)
        {
            var baseColors = new[]
            {
                new { R = 144, G = 175, B = 197 },  // mist blue
                new { R = 165, G = 196, B = 152 },  // sage green
                new { R = 210, G = 165, B = 120 },  // warm sand
                new { R = 180, G = 150, B = 180 },  // soft mauve
                new { R = 120, G = 180, B = 190 },  // teal
                new { R = 200, G = 160, B = 150 },  // rose beige
                new { R = 160, G = 185, B = 130 },  // leaf green
                new { R = 185, G = 170, B = 140 },  // light taupe
                new { R = 150, G = 170, B = 200 },  // periwinkle
                new { R = 200, G = 185, B = 150 },  // wheat
            };

            var colors = new List<Color>();

            for (int i = 0; i < colorCount; i++)
            {
                var baseColor = baseColors[i % baseColors.Length];
                int r = baseColor.R;
                int g = baseColor.G;
                int b = baseColor.B;

                if (i >= baseColors.Length)
                {
                    float brightness = 0.8f + (float)(i % baseColors.Length) /
                        baseColors.Length * 0.4f;
                    r = Math.Min((int)(r * brightness), 255);
                    g = Math.Min((int)(g * brightness), 255);
                    b = Math.Min((int)(b * brightness), 255);
                }

                colors.Add(new Color((byte)r, (byte)g, (byte)b));
            }

            return colors;
        }
    }
}
