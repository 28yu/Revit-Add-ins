using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace Tools28.Commands.BeamCommon
{
    /// <summary>
    /// 構造フレーム要素の種類（梁ラベルを付けるかの判断に使う）
    /// </summary>
    public enum FramingKind
    {
        Beam,
        /// <summary>鉛直ブレース（斜めに入るブレース・方杖）</summary>
        VerticalBrace,
        /// <summary>水平ブレース</summary>
        HorizontalBrace
    }

    /// <summary>
    /// 構造フレーム要素が梁・鉛直ブレース・水平ブレースのどれかを判定する。
    ///
    /// 判定順:
    /// 1. 構造用途（プロパティの「構造用途」）が「水平ブレース」なら水平ブレース
    /// 2. ブレースとして配置された要素（StructuralType.Brace）か、構造用途が「ブレース」「方杖」なら、
    ///    位置線の傾きで判定する（傾いていれば鉛直ブレース、ほぼ水平なら水平ブレース）
    /// 3. それ以外は梁（勾配のある梁も梁として扱う）
    /// </summary>
    public static class BeamFramingKind
    {
        // これより傾いていれば鉛直ブレースとみなす（約5°）
        private static readonly double InclinedSlope = Math.Tan(5.0 * Math.PI / 180.0);
        // 高さの差がこれ以下なら傾きに関係なく水平とみなす（10mm）
        private const double FlatToleranceFt = 10.0 / 304.8;

        public static FramingKind Classify(FamilyInstance fi)
        {
            StructuralInstanceUsage usage = StructuralInstanceUsage.Undefined;
            try { usage = fi.StructuralUsage; } catch (Exception) { }

            if (usage == StructuralInstanceUsage.HorizontalBracing)
                return FramingKind.HorizontalBrace;

            bool isBrace = fi.StructuralType == StructuralType.Brace ||
                usage == StructuralInstanceUsage.Brace ||
                usage == StructuralInstanceUsage.KickerBracing;
            if (!isBrace)
                return FramingKind.Beam;

            return IsInclined(fi) ? FramingKind.VerticalBrace : FramingKind.HorizontalBrace;
        }

        /// <summary>
        /// 梁ラベル（文字・梁タグ）を付ける要素か。鉛直ブレースには付けない
        /// </summary>
        public static bool ShouldLabel(FamilyInstance fi, bool includeHorizontalBraces)
        {
            switch (Classify(fi))
            {
                case FramingKind.VerticalBrace: return false;
                case FramingKind.HorizontalBrace: return includeHorizontalBraces;
                default: return true;
            }
        }

        private static bool IsInclined(FamilyInstance fi)
        {
            if (!(fi.Location is LocationCurve lc) || lc.Curve == null)
                return false;
            XYZ p0 = lc.Curve.GetEndPoint(0);
            XYZ p1 = lc.Curve.GetEndPoint(1);
            double dz = Math.Abs(p1.Z - p0.Z);
            if (dz <= FlatToleranceFt)
                return false;
            double horizontal = Math.Sqrt(
                (p1.X - p0.X) * (p1.X - p0.X) + (p1.Y - p0.Y) * (p1.Y - p0.Y));
            return dz > horizontal * InclinedSlope;
        }
    }
}
