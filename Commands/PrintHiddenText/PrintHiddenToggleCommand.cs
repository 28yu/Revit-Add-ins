using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.PrintHiddenText
{
    /// <summary>
    /// 設定パネルの「印刷非表示 ON/OFF」ボタン。押すたびに切り替え、ボタンの文字とアイコンを更新する。
    /// モデルは変更しない（設定ファイルに保存するだけ）。
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    public class PrintHiddenToggleCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            PrintHiddenSettings.Enabled = !PrintHiddenSettings.Enabled;
            DiagLog.Write($"PrintHiddenText ON/OFF を切り替え: {(PrintHiddenSettings.Enabled ? "ON" : "OFF")}");
            Tools28.Application.UpdatePrintHiddenButton();
            return Result.Succeeded;
        }
    }
}
