using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.LanguageSwitch
{
    [Transaction(TransactionMode.Manual)]
    public class LanguageSwitchCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            TaskDialog.Show(
                "言語切替 / Language / 语言",
                "この機能は現在実装中です。\n\nThis feature is under development.\n\n此功能正在开发中。");
            return Result.Succeeded;
        }
    }
}
