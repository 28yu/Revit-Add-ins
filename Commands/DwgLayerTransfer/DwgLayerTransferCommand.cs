using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.DwgLayerTransfer.Views;
using Tools28.Localization;

namespace Tools28.Commands.DwgLayerTransfer
{
    /// <summary>
    /// 読み込み／リンクした DWG のレイヤ表示設定（V/G の「読み込みカテゴリ」タブの内容）を、
    /// 同時に開いている別の Revit モデルから現在のモデルへ移行するコマンド。
    ///
    /// 移行先は常にアクティブなモデル。移行元は開いている他のモデルから選ぶ。
    /// 中間ファイルを介さず、開いているドキュメント同士で直接受け渡す。
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DwgLayerTransferCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData?.Application;
            Document targetDoc = uiapp?.ActiveUIDocument?.Document;

            if (targetDoc == null)
            {
                message = Loc.S("DwgVg.NoDoc");
                return Result.Cancelled;
            }

            var sourceDocs = CollectSourceDocuments(uiapp, targetDoc);
            if (sourceDocs.Count == 0)
            {
                TaskDialog.Show(Loc.S("DwgVg.Title"), Loc.S("DwgVg.NoSourceDoc"));
                return Result.Cancelled;
            }

            try
            {
                var dialog = new DwgLayerTransferDialog(targetDoc, sourceDocs);
                dialog.SetRevitOwner(commandData);
                dialog.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n\nマニュアル: https://28tools.com/addins.html";
                return Result.Failed;
            }
        }

        /// <summary>
        /// 移行元の候補となる、開いているモデルを集める。
        /// 移行先自身・リンクモデル・ファミリドキュメントは除外する。
        /// </summary>
        private static List<Document> CollectSourceDocuments(UIApplication uiapp, Document targetDoc)
        {
            var docs = new List<Document>();
            try
            {
                foreach (Document d in uiapp.Application.Documents)
                {
                    if (d == null) continue;
                    if (d.Equals(targetDoc)) continue;

                    try
                    {
                        if (d.IsLinked) continue;
                        if (d.IsFamilyDocument) continue;
                    }
                    catch { continue; }

                    docs.Add(d);
                }
            }
            catch { }
            return docs;
        }
    }
}
