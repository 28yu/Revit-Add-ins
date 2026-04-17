using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.LanguageSwitch
{
    [Transaction(TransactionMode.Manual)]
    public class AboutCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            var asm = Assembly.GetExecutingAssembly();
            var version = asm.GetName().Version;

            TaskDialog.Show(
                "28 Tools について",
                $"28 Tools\nバージョン: {version}\n\nRevit 向け業務効率化アドイン");
            return Result.Succeeded;
        }
    }
}
