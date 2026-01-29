using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace Tools28.Commands.SectionBoxCopy
{
    // 切断ボックス情報を保持する静的クラス
    public static class SectionBoxClipboard
    {
        public static BoundingBoxXYZ CopiedSectionBox { get; set; }
        public static string SourceViewName { get; set; }
        public static bool IsSectionBoxActive { get; set; }
        public static bool HasCopiedData => CopiedSectionBox != null || !IsSectionBoxActive;

        public static void Clear()
        {
            CopiedSectionBox = null;
            SourceViewName = null;
            IsSectionBoxActive = false;
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
                    message = "3Dビューをアクティブにしてから実行してください。";
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
                        message = "切断ボックスの範囲を取得できませんでした。";
                        return Result.Failed;
                    }
                }

                // 切断ボックス情報をクリップボードに保存
                SectionBoxClipboard.CopiedSectionBox = sectionBox;
                SectionBoxClipboard.SourceViewName = activeView.Name;
                SectionBoxClipboard.IsSectionBoxActive = isSectionBoxActive;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"切断ボックスのコピー中にエラーが発生しました。{ex.Message}";
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
                    message = "コピーされた切断ボックス情報がありません。先に切断ボックスをコピーしてください。";
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
                        message = "アクティブビューが3Dビューではありません。3Dビューをアクティブにするか、プロジェクトブラウザで3Dビューを選択してください。";
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
                        message = "選択された要素に3Dビューがありません。3Dビューを選択してください。";
                        return Result.Failed;
                    }
                }

                // 切断ボックスを適用
                using (Transaction trans = new Transaction(doc, "切断ボックス範囲ペースト"))
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
                                // 切断ボックスを無効にする
                                targetView.SetSectionBox(null);
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

                    // エラーがある場合のみメッセージ設定
                    if (errorCount > 0)
                    {
                        message = $"一部のビューで適用に失敗しました。成功: {successCount}件、失敗: {errorCount}件";
                    }

                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                message = $"切断ボックスのペースト中にエラーが発生しました。{ex.Message}";
                return Result.Failed;
            }
        }

        /// <summary>
        /// 3Dビューの画面更新を強制実行
        /// </summary>
        private void ForceViewUpdate(UIDocument uidoc, View3D currentActiveView, List<View3D> targetViews)
        {
            try
            {
                // 方法1: アクティブビューが対象に含まれている場合は再描画
                if (currentActiveView != null && targetViews.Contains(currentActiveView))
                {
                    // 現在のアクティブビューを再描画
                    uidoc.RefreshActiveView();
                }

                // 方法2: 各ターゲットビューを一時的にアクティブにして更新
                foreach (View3D targetView in targetViews)
                {
                    try
                    {
                        // 対象ビューを一時的にアクティブにして更新を強制
                        if (targetView.Id != uidoc.ActiveView.Id)
                        {
                            uidoc.ActiveView = targetView;
                            uidoc.RefreshActiveView();
                        }
                    }
                    catch
                    {
                        // 個別のビュー更新エラーは無視
                    }
                }

                // 方法3: 元のアクティブビューに戻す
                if (currentActiveView != null && currentActiveView.Id != uidoc.ActiveView.Id)
                {
                    try
                    {
                        uidoc.ActiveView = currentActiveView;
                        uidoc.RefreshActiveView();
                    }
                    catch
                    {
                        // 復帰エラーは無視
                    }
                }
            }
            catch
            {
                // 画面更新エラーは無視（メイン処理は成功しているため）
            }
        }
    }
}