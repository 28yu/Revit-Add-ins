using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.LanguageSwitch
{
    [Transaction(TransactionMode.Manual)]
    public class SwitchToJapaneseCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Loc.SetLanguage("JP");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class SwitchToEnglishCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Loc.SetLanguage("US");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class SwitchToChineseCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Loc.SetLanguage("CN");
            return Result.Succeeded;
        }
    }
}
