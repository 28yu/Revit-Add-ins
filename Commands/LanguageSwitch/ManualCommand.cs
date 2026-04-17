using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.LanguageSwitch
{
    [Transaction(TransactionMode.Manual)]
    public class ManualCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://28yu.github.io/28tools-manual/",
                UseShellExecute = true
            });
            return Result.Succeeded;
        }
    }
}
