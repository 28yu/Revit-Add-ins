using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.LineStyleCheck
{
    [Transaction(TransactionMode.Manual)]
    public class ExecuteLineStyleCheckCommand : IExternalCommand
    {
        private static LineStyleCheckDialog _dialog;
        private static ExternalEvent _externalEvent;
        private static LineStyleCheckEventHandler _handler;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;
                View activeView = doc.ActiveView;

                // 対象ビュータイプの確認
                if (!IsValidViewType(activeView))
                {
                    TaskDialog.Show("線種変更確認",
                        "この機能は平面図、断面図、天井伏図、3Dビューで使用できます。");
                    return Result.Failed;
                }

                // 既存のダイアログが開いている場合はアクティブにする
                if (_dialog != null && _dialog.IsLoaded && _dialog.IsVisible)
                {
                    _dialog.Activate();
                    return Result.Succeeded;
                }

                // ラインワーク変更箇所を検出
                var detector = new LineStyleOverrideDetector();
                List<OverriddenElementInfo> overriddenElements =
                    detector.FindOverriddenElements(doc, activeView);

                if (overriddenElements.Count == 0)
                {
                    TaskDialog.Show("線種変更確認",
                        "このビューにはラインワークで線種が変更された要素はありません。");
                    return Result.Succeeded;
                }

                // ハイライト表示
                using (Transaction trans = new Transaction(doc, "線種変更箇所のハイライト"))
                {
                    trans.Start();
                    ApplyHighlight(doc, activeView, overriddenElements);
                    trans.Commit();
                }

                // 外部イベントハンドラーとモードレスダイアログを作成
                _handler = new LineStyleCheckEventHandler();
                _externalEvent = ExternalEvent.Create(_handler);

                _dialog = new LineStyleCheckDialog(
                    _externalEvent,
                    _handler,
                    uidoc,
                    overriddenElements);

                // Revitのメインウインドウを親に設定
                var wih = new System.Windows.Interop.WindowInteropHelper(_dialog);
                wih.Owner = uiapp.MainWindowHandle;

                _dialog.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private bool IsValidViewType(View view)
        {
            return view.ViewType == ViewType.FloorPlan
                || view.ViewType == ViewType.CeilingPlan
                || view.ViewType == ViewType.Section
                || view.ViewType == ViewType.ThreeD
                || view.ViewType == ViewType.Elevation
                || view.ViewType == ViewType.AreaPlan;
        }

        private void ApplyHighlight(Document doc, View view,
            List<OverriddenElementInfo> overriddenElements)
        {
            OverrideGraphicSettings highlightOgs = new OverrideGraphicSettings();
            Color red = new Color(255, 0, 0);

            highlightOgs.SetProjectionLineColor(red);
            highlightOgs.SetCutLineColor(red);
            highlightOgs.SetProjectionLineWeight(5);
            highlightOgs.SetCutLineWeight(5);

            foreach (var info in overriddenElements)
            {
                view.SetElementOverrides(info.ElementId, highlightOgs);
            }
        }
    }
}
