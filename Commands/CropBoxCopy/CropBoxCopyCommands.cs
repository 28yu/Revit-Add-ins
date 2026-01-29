using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Tools28.Commands.CropBoxCopy
{
    // トリミング領域データを保持するクラス
    public class CropBoxData
    {
        public bool IsActive { get; set; }
        public bool IsVisible { get; set; }
        public BoundingBoxXYZ CropBox { get; set; }
        public string SourceViewName { get; set; }
        public ViewType SourceViewType { get; set; }
        public string ViewTypeName { get; set; }
    }

    // 静的クリップボードクラス
    public static class CropBoxClipboard
    {
        public static CropBoxData CopiedData { get; set; }
        public static string SourceName { get; set; }
        public static bool HasCopiedData => CopiedData != null;

        public static void Clear()
        {
            CopiedData = null;
            SourceName = null;
        }
    }

    // トリミング領域コピーコマンド
    [Transaction(TransactionMode.ReadOnly)]
    public class ExecuteCropBoxCopyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                // 現在のアクティブビューを取得
                View activeView = doc.ActiveView;

                if (activeView == null)
                {
                    message = "アクティブなビューがありません。";
                    return Result.Failed;
                }

                // 対象ビュータイプかチェック
                if (!IsSupportedViewType(activeView))
                {
                    message = $"このビュータイプ（{GetViewTypeName(activeView.ViewType)}）はサポートされていません。";
                    return Result.Failed;
                }

                // トリミング領域データを取得（ON/OFF関係なく）
                CropBoxData cropBoxData = new CropBoxData
                {
                    IsActive = activeView.CropBoxActive,
                    IsVisible = activeView.CropBoxActive ? activeView.CropBoxVisible : false,
                    CropBox = activeView.CropBoxActive ? activeView.CropBox : null,
                    SourceViewName = activeView.Name,
                    SourceViewType = activeView.ViewType,
                    ViewTypeName = GetViewTypeName(activeView.ViewType)
                };

                // クリップボードに保存
                CropBoxClipboard.CopiedData = cropBoxData;
                CropBoxClipboard.SourceName = activeView.Name;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"トリミング領域のコピーに失敗しました。{ex.Message}";
                return Result.Failed;
            }
        }

        private bool IsSupportedViewType(View view)
        {
            ViewType viewType = view.ViewType;
            return viewType == ViewType.FloorPlan ||
                   viewType == ViewType.CeilingPlan ||
                   viewType == ViewType.Section ||
                   viewType == ViewType.Elevation ||
                   viewType == ViewType.Detail ||
                   viewType == ViewType.ThreeD;
        }

        private string GetViewTypeName(ViewType viewType)
        {
            switch (viewType)
            {
                case ViewType.FloorPlan: return "平面図";
                case ViewType.CeilingPlan: return "天井伏図";
                case ViewType.Section: return "断面図";
                case ViewType.Elevation: return "立面図";
                case ViewType.Detail: return "詳細図";
                case ViewType.ThreeD: return "3Dビュー";
                default: return viewType.ToString();
            }
        }
    }

    // トリミング領域ペーストコマンド
    [Transaction(TransactionMode.Manual)]
    public class ExecuteCropBoxPasteCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                // コピーされたデータがあるかチェック
                if (!CropBoxClipboard.HasCopiedData)
                {
                    message = "コピーされたトリミング領域がありません。先にトリミング領域をコピーしてください。";
                    return Result.Failed;
                }

                // プロジェクトブラウザで選択されたビューを取得
                var selectedIds = uidoc.Selection.GetElementIds();
                List<View> targetViews = new List<View>();

                if (selectedIds.Count == 0)
                {
                    // 選択がない場合：アクティブビューにペースト
                    View activeView = doc.ActiveView;

                    if (activeView == null)
                    {
                        message = "アクティブなビューがありません。";
                        return Result.Failed;
                    }

                    if (!IsSupportedViewType(activeView))
                    {
                        message = $"このビュータイプ（{GetViewTypeName(activeView.ViewType)}）はサポートされていません。";
                        return Result.Failed;
                    }

                    targetViews.Add(activeView);
                }
                else
                {
                    // 選択がある場合：選択されたビューにペースト
                    foreach (ElementId id in selectedIds)
                    {
                        Element element = doc.GetElement(id);
                        if (element is View view && IsSupportedViewType(view))
                        {
                            targetViews.Add(view);
                        }
                    }

                    if (targetViews.Count == 0)
                    {
                        message = "対応するビューが選択されていません。";
                        return Result.Failed;
                    }
                }

                // トランザクション開始
                using (Transaction trans = new Transaction(doc, "トリミング領域の適用"))
                {
                    trans.Start();

                    CropBoxData copiedData = CropBoxClipboard.CopiedData;
                    int successCount = 0;
                    int failCount = 0;

                    foreach (View targetView in targetViews)
                    {
                        try
                        {
                            if (copiedData.IsActive)
                            {
                                // コピー元がトリミングON → ペースト先もONにして領域をコピー

                                // まずトリミングをONにする
                                targetView.CropBoxActive = true;

                                // トリミング領域をコピー
                                if (copiedData.CropBox != null)
                                {
                                    targetView.CropBox = copiedData.CropBox;
                                }

                                // 表示/非表示設定をコピー
                                targetView.CropBoxVisible = copiedData.IsVisible;
                            }
                            else
                            {
                                // コピー元がトリミングOFF → ペースト先もOFFにする
                                targetView.CropBoxActive = false;
                            }

                            successCount++;
                        }
                        catch
                        {
                            failCount++;
                        }
                    }

                    trans.Commit();

                    // エラーがある場合のみメッセージ設定
                    if (failCount > 0)
                    {
                        message = $"一部のビューで適用に失敗しました。成功: {successCount}件、失敗: {failCount}件";
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"トリミング領域の適用に失敗しました。{ex.Message}";
                return Result.Failed;
            }
        }

        private bool IsSupportedViewType(View view)
        {
            ViewType viewType = view.ViewType;
            return viewType == ViewType.FloorPlan ||
                   viewType == ViewType.CeilingPlan ||
                   viewType == ViewType.Section ||
                   viewType == ViewType.Elevation ||
                   viewType == ViewType.Detail ||
                   viewType == ViewType.ThreeD;
        }

        private string GetViewTypeName(ViewType viewType)
        {
            switch (viewType)
            {
                case ViewType.FloorPlan: return "平面図";
                case ViewType.CeilingPlan: return "天井伏図";
                case ViewType.Section: return "断面図";
                case ViewType.Elevation: return "立面図";
                case ViewType.Detail: return "詳細図";
                case ViewType.ThreeD: return "3Dビュー";
                default: return viewType.ToString();
            }
        }
    }
}