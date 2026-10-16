using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Lighting;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// 照明器具のファミリの「測光」の設定を読み、描画プログラムの光源にする。
    ///
    /// 明るさは光度（カンデラ）に直してから 20000 で割る（描画プログラムでは太陽 5 = 10 万ルクスなので、1 = 2 万ルクス）。
    /// 光束（ルーメン）の場合は、配光（全方向・半球・スポット）の広がり（立体角）で割って光度にする。
    /// </summary>
    internal static class LightReader
    {
        private const double FeetToMeters = 0.3048;
        private const double LuxPerUnit = 20000.0;

        /// <param name="transform">光源の位置・向き（プロジェクト座標。フィート）</param>
        public static RenderLight Read(Document doc, ElementId elementId, Transform transform)
        {
            if (doc == null || elementId == null || elementId == ElementId.InvalidElementId) return null;
            LightType type = LightType.GetLightTypeFromInstance(doc, elementId);
            if (type == null) return null;

            var light = new RenderLight();
            XYZ p = transform.Origin;
            light.Position = new[] { (float)(p.X * FeetToMeters), (float)(p.Y * FeetToMeters), (float)(p.Z * FeetToMeters) };
            // Revit の光源は、ファミリの下向き（-Z）に光る
            XYZ down = transform.BasisZ.Negate();
            if (down.GetLength() > 1e-9) down = down.Normalize();
            light.Direction = new[] { (float)down.X, (float)down.Y, (float)down.Z };

            // 配光と、光が広がる立体角
            double solidAngle = 4 * Math.PI;
            switch (type.GetLightDistribution())
            {
                case SpotLightDistribution spot:
                    light.Type = 2;
                    double field = Math.Max(spot.SpotFieldAngle, 0.01), beam = Math.Min(spot.SpotBeamAngle, field);
                    light.CosOuter = (float)Math.Cos(field / 2);
                    light.CosInner = (float)Math.Cos(beam / 2);
                    solidAngle = 2 * Math.PI * (1 - Math.Cos(field / 2));
                    break;
                case HemisphericalLightDistribution _:
                case PhotometricWebLightDistribution _:  // 配光データ（IES）は読まず、下向きの半球とみなす
                    light.Type = 1;
                    solidAngle = 2 * Math.PI;
                    break;
                default:
                    light.Type = 0;
                    break;
            }

            // 明るさ → 光度（カンデラ）
            double candela = 0;
            switch (type.GetInitialIntensity())
            {
                case InitialFluxIntensity flux:
                    candela = UnitUtils.ConvertFromInternalUnits(flux.Flux, UnitTypeId.Lumens) / solidAngle;
                    break;
                case InitialWattageIntensity watt:
                    double lumens = UnitUtils.ConvertFromInternalUnits(watt.Wattage, UnitTypeId.Watts) *
                                    UnitUtils.ConvertFromInternalUnits(watt.Efficacy, UnitTypeId.LumensPerWatt);
                    candela = lumens / solidAngle;
                    break;
                case InitialLuminousIntensity luminous:
                    candela = UnitUtils.ConvertFromInternalUnits(luminous.Luminosity, UnitTypeId.Candelas);
                    break;
                case InitialIlluminanceIntensity illuminance:
                    double d = illuminance.Distance * FeetToMeters;
                    candela = UnitUtils.ConvertFromInternalUnits(illuminance.Illuminance, UnitTypeId.Lux) * d * d;
                    break;
            }
            double loss = type.GetLossFactor()?.LossFactorValue ?? 1.0;
            if (loss <= 0 || loss > 1) loss = 1.0;
            candela *= loss;
            if (!(candela > 0)) return null;

            // 色 = 色温度の色 × カラーフィルタ（明るさは光度で決めるので、色の明るさは 1 にそろえる）
            double kelvin = type.GetInitialColor()?.TemperatureValue ?? 3000;
            double[] rgb = KelvinToLinearRgb(kelvin);
            Color filter = type.ColorFilter;
            if (filter != null && filter.IsValid)
            {
                rgb[0] *= SrgbToLinear(filter.Red / 255.0);
                rgb[1] *= SrgbToLinear(filter.Green / 255.0);
                rgb[2] *= SrgbToLinear(filter.Blue / 255.0);
            }
            double luminance = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2];
            if (luminance <= 1e-6) return null;
            double scale = candela / LuxPerUnit / luminance;
            light.Intensity = new[] { (float)(rgb[0] * scale), (float)(rgb[1] * scale), (float)(rgb[2] * scale) };

            // 光る部分の大きさ（影のぼけ方）
            double size = 0.06;
            switch (type.GetLightShape())
            {
                case CircleLightShape circle:
                    size = circle.EmitDiameter * FeetToMeters;
                    break;
                case RectangleLightShape rect:
                    size = Math.Sqrt(rect.EmitLength * rect.EmitWidth) * FeetToMeters;
                    break;
                case LineLightShape line:
                    size = line.EmitLength * FeetToMeters * 0.5;
                    break;
            }
            light.Radius = (float)Math.Max(0.01, Math.Min(0.5, size / 2));
            return light;
        }

        /// <summary>色温度（ケルビン）の光の色（リニア）。黒体放射の色を近似する式（Tanner Helland の式）。</summary>
        internal static double[] KelvinToLinearRgb(double kelvin)
        {
            double t = Math.Max(1000, Math.Min(40000, kelvin)) / 100.0;
            double r, g, b;
            if (t <= 66)
            {
                r = 255;
                g = 99.4708025861 * Math.Log(t) - 161.1195681661;
                b = t <= 19 ? 0 : 138.5177312231 * Math.Log(t - 10) - 305.0447927307;
            }
            else
            {
                r = 329.698727446 * Math.Pow(t - 60, -0.1332047592);
                g = 288.1221695283 * Math.Pow(t - 60, -0.0755148492);
                b = 255;
            }
            return new[]
            {
                SrgbToLinear(Math.Max(0, Math.Min(255, r)) / 255.0),
                SrgbToLinear(Math.Max(0, Math.Min(255, g)) / 255.0),
                SrgbToLinear(Math.Max(0, Math.Min(255, b)) / 255.0),
            };
        }

        private static double SrgbToLinear(double s) => s <= 0.04045 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);
    }
}
