using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// Revit 2024 で ElementId は long 化、Revit 2026 では IntegerValue プロパティが
    /// 削除された。両バージョンを跨いで int 値が取得できるよう拡張メソッドを提供する。
    /// </summary>
    internal static class ElementIdHelper
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
