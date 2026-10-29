using Autodesk.Revit.DB;

namespace Tools28
{
    /// <summary>
    /// Revit API のバージョン差異を吸収する共通ヘルパー。
    /// Revit 2026 で ElementId.IntegerValue が削除され Value (long) に変わったため、
    /// IntValue() 拡張メソッドで両バージョンを跨いで int 値を取得できるようにする。
    /// </summary>
    internal static class RevitCompatibility
    {
        public static int IntValue(this ElementId id)
        {
#if REVIT2026
            return (int)id.Value;
#else
            return id.IntegerValue;
#endif
        }

        /// <summary>要素 ID の値（Revit 2024 以降は 64 ビット。描画プログラムへ要素を伝えるときに使う）。</summary>
        public static long LongValue(this ElementId id)
        {
#if REVIT2021 || REVIT2022 || REVIT2023
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }
    }
}
