using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace Tools28.Commands.ViewCopy
{
    // 視点情報を保存するための静的クラス
    public static class ViewOrientationClipboard
    {
        public static ViewOrientation3D CopiedOrientation { get; set; }
        public static string SourceViewName { get; set; }
        public static bool HasCopiedOrientation => CopiedOrientation != null;
    }

    // 視点コピーコマンド
    [Transaction(TransactionMode.Manual)]
    public class ExecuteViewCopyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // 現在のアクティブビューが3Dビューかチェック
                View activeView = doc.ActiveView;
                if (!(activeView is View3D sourceView))
                {
                    message = "現在のアクティブビューが3Dビューではありません。3Dビューをアクティブにしてから実行してください。";
                    return Result.Failed;
                }

                // 視点情報をクリップボードに保存
                ViewOrientationClipboard.CopiedOrientation = sourceView.GetOrientation();
                ViewOrientationClipboard.SourceViewName = sourceView.Name;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"視点のコピー中にエラーが発生しました: {ex.Message}";
                return Result.Failed;
            }
        }
    }

    // 視点ペーストコマンド
    [Transaction(TransactionMode.Manual)]
    public class ExecuteViewPasteCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // コピーされた視点情報があるかチェック
                if (!ViewOrientationClipboard.HasCopiedOrientation)
                {
                    message = "コピーされた視点情報がありません。先に3Dビューの視点をコピーしてください。";
                    return Result.Failed;
                }

                // プロジェクトブラウザで選択されている3Dビューを取得
                var selectedElementIds = uidoc.Selection.GetElementIds();
                List<View3D> targetViews = new List<View3D>();
                View3D currentActiveView = null;

                if (selectedElementIds.Count == 0)
                {
                    // 選択がない場合：アクティブビューにペースト
                    View activeView = doc.ActiveView;

                    if (!(activeView is View3D activeView3D))
                    {
                        message = "アクティブビューが3Dビューではありません。3Dビューをアクティブにするか、プロジェクトブラウザで3Dビューを選択してください。";
                        return Result.Failed;
                    }

                    // テンプレートビューは除外
                    if (activeView3D.IsTemplate)
                    {
                        message = "テンプレートビューには視点を適用できません。";
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

                    foreach (ElementId id in selectedElementIds)
                    {
                        Element element = doc.GetElement(id);

                        if (element is View3D targetView)
                        {
                            // テンプレートビューは除外
                            if (!targetView.IsTemplate)
                            {
                                targetViews.Add(targetView);
                            }
                        }
                    }

                    // 有効な3Dビューがない場合
                    if (targetViews.Count == 0)
                    {
                        message = "有効な3Dビューが選択されていません。テンプレートビュー以外の3Dビューを選択してください。";
                        return Result.Failed;
                    }
                }

                // 視点を一括ペースト
                using (Transaction trans = new Transaction(doc, "3Dビュー視点一括ペースト"))
                {
                    trans.Start();

                    int successCount = 0;
                    int errorCount = 0;

                    // 各ターゲットビューに視点を適用
                    foreach (View3D targetView in targetViews)
                    {
                        try
                        {
                            targetView.SetOrientation(ViewOrientationClipboard.CopiedOrientation);
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
                message = $"視点の適用中にエラーが発生しました: {ex.Message}";
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