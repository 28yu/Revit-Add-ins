using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Tools28.Commands.BeamCommon
{
    /// <summary>色分けする梁の範囲（水平・斜め）</summary>
    public enum BeamSlopeFilter
    {
        /// <summary>水平の梁と斜めの梁の両方</summary>
        All,
        /// <summary>水平の梁だけ</summary>
        FlatOnly,
        /// <summary>斜めの梁だけ</summary>
        SlopedOnly
    }

    /// <summary>設定画面に出す、水平・斜めごとの本数</summary>
    public class BeamSlopeCounts
    {
        public int All { get; set; }
        public int Flat { get; set; }
        public int Sloped { get; set; }

        public int CountOf(BeamSlopeFilter filter)
        {
            switch (filter)
            {
                case BeamSlopeFilter.FlatOnly: return Flat;
                case BeamSlopeFilter.SlopedOnly: return Sloped;
                default: return All;
            }
        }
    }

    /// <summary>
    /// 梁天端・梁下端色分けで、どの梁を色分けするかを決める。
    ///
    /// - 参照レベル: ビューに映っている梁には、ビュー範囲によって別の階の梁も混ざる。
    ///   梁の参照レベルで、その階の梁だけに絞る（別の階の梁の値・凡例・ラベルに触らない）
    /// - 水平・斜め: 位置線の始点と終点の高さが違う梁を「斜め」とする
    /// </summary>
    public static class BeamSelection
    {
        // 始点と終点の高さの差がこれ以下なら水平とみなす（1mm。丸めの誤差を吸収する）
        private const double FlatToleranceFt = 1.0 / 304.8;

        /// <summary>
        /// 梁の参照レベル。取得できなければ InvalidElementId
        /// （その場合は絞り込みで除外しない＝従来どおり対象にする）
        /// </summary>
        public static ElementId ReferenceLevelId(FamilyInstance fi)
        {
            try
            {
                Parameter p = fi.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                if (p != null && p.StorageType == StorageType.ElementId)
                {
                    ElementId id = p.AsElementId();
                    if (id != null && id != ElementId.InvalidElementId)
                        return id;
                }
            }
            catch (Exception) { }

            try
            {
                if (fi.LevelId != null && fi.LevelId != ElementId.InvalidElementId)
                    return fi.LevelId;
            }
            catch (Exception) { }

            try
            {
                if (fi.Host is Level hostLevel)
                    return hostLevel.Id;
            }
            catch (Exception) { }

            return ElementId.InvalidElementId;
        }

        /// <summary>参照レベルが指定のレベルの梁か（参照レベルが分からない梁も含める）</summary>
        public static bool IsOnLevel(FamilyInstance fi, ElementId levelId)
        {
            ElementId refId = ReferenceLevelId(fi);
            return refId == ElementId.InvalidElementId || refId == levelId;
        }

        /// <summary>始点と終点の高さが違う（斜めの）梁か</summary>
        public static bool IsSloped(FamilyInstance fi)
        {
            if (!(fi.Location is LocationCurve lc) || lc.Curve == null)
                return false;
            try
            {
                double dz = Math.Abs(lc.Curve.GetEndPoint(1).Z - lc.Curve.GetEndPoint(0).Z);
                return dz > FlatToleranceFt;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 選んだ範囲に入る梁か。
        /// 「斜めの梁だけ」には鉛直ブレース（斜めに入るブレース・方杖）は含めない（梁ではないため）
        /// </summary>
        public static bool Matches(FamilyInstance fi, BeamSlopeFilter filter)
        {
            switch (filter)
            {
                case BeamSlopeFilter.FlatOnly:
                    return !IsSloped(fi);
                case BeamSlopeFilter.SlopedOnly:
                    return IsSloped(fi) && BeamFramingKind.Classify(fi) == FramingKind.Beam;
                default:
                    return true;
            }
        }

        /// <summary>水平・斜めごとの本数を数える</summary>
        public static BeamSlopeCounts Count(IEnumerable<FamilyInstance> beams)
        {
            var list = beams.ToList();
            return new BeamSlopeCounts
            {
                All = list.Count,
                Flat = list.Count(b => Matches(b, BeamSlopeFilter.FlatOnly)),
                Sloped = list.Count(b => Matches(b, BeamSlopeFilter.SlopedOnly))
            };
        }
    }
}
