using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Licensing;
using Tools28.Localization;

namespace Tools28.Commands.LanguageSwitch
{
    [Transaction(TransactionMode.Manual)]
    public class AboutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DiagLog.Cmd("About", "Execute 開始");
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                var expiry = ExpiryManager.ExpiryDate.ToString("yyyy/MM/dd");
                var status = ExpiryManager.IsExpired
                    ? Loc.S("About.ExpiredSuffix")
                    : string.Format(Loc.S("About.RemainingSuffix"), ExpiryManager.DaysRemaining);

                DiagLog.Cmd("About", "TaskDialog.Show 直前");
                TaskDialog.Show(
                    Loc.S("About.Title"),
                    string.Format(Loc.S("About.Message"), version, expiry, status));
                DiagLog.Cmd("About", "TaskDialog.Show 戻り");

                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                DiagLog.Cmd("About", $"例外: {ex}");
                throw;
            }
            finally
            {
                DiagLog.Cmd("About", "Execute 終了");
            }
        }
    }
}
