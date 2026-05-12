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
    }
}
