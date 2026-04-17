using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Localization;

namespace Tools28.Commands.LanguageSwitch
{
    [Transaction(TransactionMode.Manual)]
    public class AboutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            TaskDialog.Show(Loc.S("About.Title"), string.Format(Loc.S("About.Message"), version));
            return Result.Succeeded;
        }
    }
}
