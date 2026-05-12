using Autodesk.Revit.DB;

namespace Tools28.Commands.FormworkCalculator.Engine
{
    /// <summary>
    /// FormworkCalculator 内で Tools28.RevitCompatibility.IntValue() へアクセスできるよう
    /// 同名の拡張メソッドを中継する。実装は RevitCompatibility.cs に集約。
    /// </summary>
    internal static class ElementIdHelper
    {
        public static int IntValue(this ElementId id) => RevitCompatibility.IntValue(id);
    }
}

