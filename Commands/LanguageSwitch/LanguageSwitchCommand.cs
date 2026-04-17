using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.LanguageSwitch
{
    [Transaction(TransactionMode.Manual)]
    public class SwitchToJapaneseCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            TaskDialog.Show("言語切替", "日本語に切り替えます。\n（この機能は現在実装中です）");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class SwitchToEnglishCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            TaskDialog.Show("Language", "Switching to English.\n(This feature is under development)");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class SwitchToChineseCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            TaskDialog.Show("语言", "切换为中文。\n（此功能正在开发中）");
            return Result.Succeeded;
        }
    }
}
