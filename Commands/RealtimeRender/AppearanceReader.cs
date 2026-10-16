using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Visual;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// マテリアルの「外観」タブの内容（Revit のレンダリング用の情報 = Asset）から、描画に使う値を読み取ったもの。
    /// 読めなかった項目は null のまま（そのときは Revit の画面表示用の色などを使う）。
    /// </summary>
    internal class AppearanceInfo
    {
        /// <summary>色（sRGB, 0〜1）。</summary>
        public double[] Color;
        /// <summary>色のテクスチャ画像（外観の「イメージ」）。</summary>
        public BitmapInfo ColorImage;
        /// <summary>イメージと色の混ぜ具合（1 = イメージだけ）。Revit の「イメージ フェード」。</summary>
        public double ImageFade = 1.0;
        /// <summary>粗さ（0 = つるつる〜1 = ざらざら）。</summary>
        public double? Roughness;
        /// <summary>粗さがテクスチャ画像のとき（画像の平均を粗さとして使う）。</summary>
        public BitmapInfo RoughnessImage;
        /// <summary>正面から見たときの反射率（外観の「反射率」）。</summary>
        public double? F0;
        /// <summary>透明度（0 = 不透明〜1 = 透明）。</summary>
        public double? Transparency;
        public bool Metal;
        public bool Glass;
        public bool Mirror;
    }

    /// <summary>テクスチャ画像のファイルと、貼り方（実寸の大きさ・ずらし・回転）。長さはメートル。</summary>
    internal class BitmapInfo
    {
        public string Paths = "";
        public double ScaleX = 1.0, ScaleY = 1.0;
        public double OffsetX, OffsetY;
        public double AngleDegrees;
        /// <summary>明るさの倍率（Revit の「明度」）。</summary>
        public double Brightness = 1.0;
        public bool Invert;
        /// <summary>色付け（sRGB, 0〜1）。null なら色付けしない。</summary>
        public double[] Tint;

        /// <summary>同じ画像・同じ色付けなら同じテクスチャを使い回すための目印。</summary>
        public string BakeKey => $"{Paths}|{Brightness:F3}|{Invert}|{FormatColor(Tint)}";

        /// <summary>
        /// 形の面上の位置（Revit の UV。フィート）を、テクスチャ画像上の位置（1 = 画像 1 枚分）に直す。
        /// </summary>
        public void Map(double uFeet, double vFeet, out float u, out float v)
        {
            double x = uFeet * 0.3048 - OffsetX, y = vFeet * 0.3048 - OffsetY;
            double a = -AngleDegrees * Math.PI / 180.0;
            double rx = x * Math.Cos(a) - y * Math.Sin(a);
            double ry = x * Math.Sin(a) + y * Math.Cos(a);
            u = (float)(rx / ScaleX);
            v = (float)(ry / ScaleY);
        }

        internal static string FormatColor(double[] c) => c == null ? "-" : $"{c[0]:F3},{c[1]:F3},{c[2]:F3}";
    }

    /// <summary>
    /// Revit の外観（Asset）を読む。外観の種類（スキーマ）によってプロパティ名が違うので、
    /// 名前で見分けて、色・イメージ・粗さ・反射率・透明度・金属かどうかを取り出す。
    /// </summary>
    internal static class AppearanceReader
    {
        // 色（またはイメージ）を表すプロパティ名。前にあるものほど優先する
        private static readonly string[] ColorNames =
        {
            "generic_diffuse",            // 一般
            "opaque_albedo",              // 不透明（新しい物理ベースの外観。コンクリート 現場打ち など）
            "metal_f0",                   // 金属（新しい物理ベースの外観）
            "layered_diffuse",            // 塗装（重ね塗り）
            "transparent_color",          // 透明（新しい物理ベースの外観）
            "ceramic_color", "concrete_color", "hardwood_color", "masonrycmu_color", "metal_color",
            "plasticvinyl_color", "stone_color", "wallpaint_color",
            "glazing_transmittance_map", "solidglass_transmittance_custom_color", "water_tint_color",
            "mirror_tintcolor",
        };

        public static AppearanceInfo Read(Asset asset)
        {
            if (asset == null) return null;
            var info = new AppearanceInfo();
            var props = new Dictionary<string, AssetProperty>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < asset.Size; i++)
            {
                AssetProperty p = asset.Get(i);
                if (p != null && !string.IsNullOrEmpty(p.Name) && !props.ContainsKey(p.Name))
                    props[p.Name] = p;
            }

            foreach (string name in ColorNames)
            {
                if (!props.TryGetValue(name, out AssetProperty p)) continue;
                info.Color = ReadColor(p);
                info.ColorImage = ReadBitmap(p);
                if (info.Color != null || info.ColorImage != null) break;
            }
            if (props.TryGetValue("generic_diffuse_image_fade", out AssetProperty fade))
                info.ImageFade = Clamp01(ReadDouble(fade) ?? 1.0);

            // 粗さ: 新しい外観は「粗さ」、一般は「光沢」（粗さ = 1 − 光沢）
            if (props.TryGetValue("surface_roughness", out AssetProperty rough))
            {
                info.Roughness = ReadDouble(rough);
                info.RoughnessImage = ReadBitmap(rough);
            }
            else if (props.TryGetValue("generic_glossiness", out AssetProperty gloss))
            {
                double? g = ReadDouble(gloss);
                if (g.HasValue) info.Roughness = 1.0 - Clamp01(g.Value);
            }

            if (props.TryGetValue("opaque_f0", out AssetProperty f0))
                info.F0 = ReadDouble(f0);
            else if (props.TryGetValue("generic_reflectivity_at_0deg", out AssetProperty refl))
                info.F0 = ReadDouble(refl);

            if (props.TryGetValue("generic_transparency", out AssetProperty transparency))
                info.Transparency = ReadDouble(transparency);

            bool HasPrefix(string prefix)
            {
                foreach (string key in props.Keys)
                    if (key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            }

            info.Metal = HasPrefix("metal_") ||
                         (props.TryGetValue("generic_is_metal", out AssetProperty isMetal) && isMetal is AssetPropertyBoolean b && b.Value);
            info.Mirror = HasPrefix("mirror_");
            info.Glass = HasPrefix("glazing_") || HasPrefix("solidglass_") || HasPrefix("transparent_") || HasPrefix("water_");
            return info;
        }

        /// <summary>色のプロパティを sRGB（0〜1）で読む。色の代わりに画像がつながっていることもある。</summary>
        private static double[] ReadColor(AssetProperty p)
        {
            if (p is AssetPropertyDoubleArray4d c4)
            {
                IList<double> v = c4.GetValueAsDoubles();
                if (v != null && v.Count >= 3) return new[] { Clamp01(v[0]), Clamp01(v[1]), Clamp01(v[2]) };
            }
            else if (p is AssetPropertyDoubleArray3d c3)
            {
                IList<double> v = c3.GetValueAsDoubles();
                if (v != null && v.Count >= 3) return new[] { Clamp01(v[0]), Clamp01(v[1]), Clamp01(v[2]) };
            }
            return null;
        }

        private static double? ReadDouble(AssetProperty p)
        {
            switch (p)
            {
                case AssetPropertyDouble d: return d.Value;
                case AssetPropertyFloat f: return f.Value;
                case AssetPropertyInteger i: return i.Value;
                default: return null;
            }
        }

        /// <summary>プロパティにつながっているテクスチャ画像（UnifiedBitmap）を読む。無ければ null。</summary>
        private static BitmapInfo ReadBitmap(AssetProperty p)
        {
            Asset bitmap;
            try
            {
                if (p.NumberOfConnectedProperties == 0) return null;
                bitmap = p.GetSingleConnectedAsset();
            }
            catch
            {
                return null;
            }
            if (bitmap == null) return null;

            var info = new BitmapInfo();
            if (bitmap.FindByName(UnifiedBitmap.UnifiedbitmapBitmap) is AssetPropertyString path)
                info.Paths = path.Value ?? "";
            if (string.IsNullOrWhiteSpace(info.Paths)) return null;

            info.ScaleX = Distance(bitmap, UnifiedBitmap.TextureRealWorldScaleX) ?? info.ScaleX;
            info.ScaleY = Distance(bitmap, UnifiedBitmap.TextureRealWorldScaleY) ?? info.ScaleY;
            if (info.ScaleX < 1e-6) info.ScaleX = 1.0;
            if (info.ScaleY < 1e-6) info.ScaleY = 1.0;
            info.OffsetX = Distance(bitmap, UnifiedBitmap.TextureRealWorldOffsetX) ?? 0.0;
            info.OffsetY = Distance(bitmap, UnifiedBitmap.TextureRealWorldOffsetY) ?? 0.0;
            if (bitmap.FindByName(UnifiedBitmap.TextureWAngle) is AssetProperty angle)
                info.AngleDegrees = ReadDouble(angle) ?? 0.0;
            if (bitmap.FindByName(UnifiedBitmap.UnifiedbitmapRGBAmount) is AssetProperty amount)
                info.Brightness = ReadDouble(amount) ?? 1.0;
            if (bitmap.FindByName(UnifiedBitmap.UnifiedbitmapInvert) is AssetPropertyBoolean invert)
                info.Invert = invert.Value;
            if (bitmap.FindByName("common_Tint_toggle") is AssetPropertyBoolean tintOn && tintOn.Value &&
                bitmap.FindByName("common_Tint_color") is AssetProperty tint)
                info.Tint = ReadColor(tint);
            return info;
        }

        /// <summary>長さのプロパティをメートルで読む（Revit はプロパティごとに単位を持っている）。</summary>
        private static double? Distance(Asset asset, string name)
        {
            switch (asset.FindByName(name))
            {
                case AssetPropertyDistance d:
                    try
                    {
                        return UnitUtils.ConvertToInternalUnits(d.Value, d.GetUnitTypeId()) * 0.3048;
                    }
                    catch
                    {
                        return d.Value * 0.0254;  // 単位が読めないときはインチとみなす（Revit の既定）
                    }
                case AssetPropertyDouble v:
                    return v.Value * 0.0254;
                default:
                    return null;
            }
        }

        private static double Clamp01(double v) => v < 0 ? 0 : v > 1 ? 1 : v;
    }
}
