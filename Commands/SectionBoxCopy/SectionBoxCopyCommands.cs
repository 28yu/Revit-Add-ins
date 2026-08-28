using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Tools28.Localization;

namespace Tools28.Commands.SectionBoxCopy
{
    // 切断ボックス情報を保持する静的クラス
    public static class SectionBoxClipboard
    {
        public static BoundingBoxXYZ CopiedSectionBox { get; set; }
        public static string SourceViewName { get; set; }

        /// <summary>コピー元で切断ボックスが有効だったか。</summary>
        public static bool IsSectionBoxActive { get; set; }

        /// <summary>
        /// コピーを1度でも実行したか。
        /// 「まだコピーしていない」と「コピー元が切断ボックスOFFだった」は
        /// どちらも CopiedSectionBox == null になるため、この独立したフラグで区別する。
        /// （旧実装は !IsSectionBoxActive を有効なコピーとみなしていたため、
        ///   起動直後にペーストすると対象ビューの切断ボックスが黙って解除されていた）
        /// </summary>
        public static bool HasCopied { get; set; }

        public static bool HasCopiedData => HasCopied;

        public static void Clear()
        {
            CopiedSectionBox = null;
            SourceViewName = null;
            IsSectionBoxActive = false;
            HasCopied = false;
        }
    }

    // 切断ボックスコピーコマンド
    [Transaction(TransactionMode.Manual)]
    public class ExecuteSectionBoxCopyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                View activeView = doc.ActiveView;

                // アクティブビューが3Dビューかチェック
                if (!(activeView is View3D view3d))
                {
                    message = Loc.S("Common.Need3DViewActive");
                    return Result.Failed;
                }

                // 切断ボックスの状態を確認
                bool isSectionBoxActive = view3d.IsSectionBoxActive;
                BoundingBoxXYZ sectionBox = null;

                if (isSectionBoxActive)
                {
                    sectionBox = view3d.GetSectionBox();
                    if (sectionBox == null)
                    {
                        message = Loc.S("SectionBox.NoRange");
                        return Result.Failed;
                    }
                }

                // 切断ボックス情報をクリップボードに保存
                SectionBoxClipboard.CopiedSectionBox = sectionBox;
                SectionBoxClipboard.SourceViewName = activeView.Name;
                SectionBoxClipboard.IsSectionBoxActive = isSectionBoxActive;
                SectionBoxClipboard.HasCopied = true;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("SectionBox.CopyError"), ex.Message);
                return Result.Failed;
            }
        }
    }

    // 切断ボックスペーストコマンド
    [Transaction(TransactionMode.Manual)]
    public class ExecuteSectionBoxPasteCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // コピーされた切断ボックス情報があるかチェック
                if (!SectionBoxClipboard.HasCopiedData)
                {
                    message = Loc.S("SectionBox.NothingCopied");
                    return Result.Failed;
                }

                // プロジェクトブラウザで選択されたビューを取得
                var selectedIds = uidoc.Selection.GetElementIds();
                List<View3D> targetViews = new List<View3D>();
                View3D currentActiveView = null;

                if (selectedIds.Count == 0)
                {
                    // 選択がない場合：アクティブビューにペースト
                    View activeView = doc.ActiveView;

                    if (!(activeView is View3D activeView3D))
                    {
                        message = Loc.S("Common.Need3DViewOrSelect");
                        return Result.Failed;
                    }

                    targetViews.Add(activeView3D);
                    currentActiveView = activeView3D; // アクティブビューを記録
                }
                else
                {
                    // 選択がある場合：選択された3Dビューにペースト
                    View activeView = doc.ActiveView;
                    if (activeView is View3D activeView3D)
                    {
                        currentActiveView = activeView3D; // 現在のアクティブビューを記録
                    }

                    foreach (ElementId id in selectedIds)
                    {
                        Element element = doc.GetElement(id);
                        if (element is View3D view3d)
                        {
                            targetViews.Add(view3d);
                        }
                    }

                    if (targetViews.Count == 0)
                    {
                        message = Loc.S("SectionBox.NoView3DSelected");
                        return Result.Failed;
                    }
                }

                // 切断ボックスを適用
                using (Transaction trans = new Transaction(doc, Loc.S("SectionBox.Txn.Paste")))
                {
                    trans.Start();

                    int successCount = 0;
                    int errorCount = 0;

                    foreach (View3D targetView in targetViews)
                    {
                        try
                        {
                            if (SectionBoxClipboard.IsSectionBoxActive && SectionBoxClipboard.CopiedSectionBox != null)
                            {
                                // 切断ボックスを有効にして範囲を設定
                                targetView.SetSectionBox(SectionBoxClipboard.CopiedSectionBox);
                            }
                            else
                            {
                                // 切断ボックスを無効にする。
                                // SetSectionBox(null) は Revit のバージョンによって例外になるため、
                                // 明示的にプロパティで OFF にする。
                                targetView.IsSectionBoxActive = false;
                            }

                            successCount++;
                        }
                        catch
                        {
                            errorCount++;
                        }
                    }

                    trans.Commit();

                    // 画面更新を強制実行
                    ForceViewUpdate(uidoc, currentActiveView, targetViews);

                    // Revit は Result.Succeeded のとき message を表示しないため、
                    // 部分失敗は TaskDialog で明示的に伝える。
                    if (errorCount > 0)
                    {
                        TaskDialog.Show(Loc.S("Common.Warning"),
                            string.Format(Loc.S("Common.PartialFail"), successCount, errorCount));
                    }

                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("SectionBox.PasteError"), ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// 画面更新。
        /// 非アクティブなビューは次に開いた時点で最新状態で描画されるため、
        /// 各ビューを一時的にアクティブ化する必要はない。
        /// （旧実装は対象ビューを1つずつアクティブにしていたため、
        ///   多数のビューへペーストすると全ビューが開き著しく遅くなっていた）
        /// </summary>
        private void ForceViewUpdate(UIDocument uidoc, View3D currentActiveView, List<View3D> targetViews)
        {
            try
            {
                if (currentActiveView != null &&
                    targetViews.Any(v => v != null && v.Id == currentActiveView.Id))
                {
                    uidoc.RefreshActiveView();
                }
            }
            catch
            {
                // 画面更新エラーは無視（メイン処理は成功しているため）
            }
        }
    }
}