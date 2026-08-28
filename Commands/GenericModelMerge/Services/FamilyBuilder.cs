using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace Tools28.Commands.GenericModelMerge.Services
{
    /// <summary>
    /// まとめた形状を一般モデルファミリ (.rfa) として作成し、プロジェクトへロード・配置する。
    ///
    /// 手順は 2 段階に分かれる:
    ///   Phase A (CreateFamilyFile): ファミリ文書を新規作成 → FreeFormElement で形状を作成
    ///                               → 材質を設定 → .rfa として保存・クローズ
    ///   Phase B (LoadAndPlace):     プロジェクトのトランザクション内でロードして配置
    ///
    /// Document.SaveAs は「どこかにトランザクションが開いていると失敗する」ため、
    /// ファミリ文書の作成・保存は必ずプロジェクトのトランザクションを開く前に済ませる。
    ///
    /// 形状の座標はプロジェクトの座標のまま持ち込むので、原点に配置すれば
    /// 元の要素と同じ位置に重なる。
    /// </summary>
    internal static class FamilyBuilder
    {
        /// <summary>Phase A の結果。</summary>
        internal class FamilyFileOutcome
        {
            public bool Success { get; set; }
            public string SavedPath { get; set; }
            /// <summary>FreeFormElement の作成に失敗した形状の数。</summary>
            public int FailedSolidCount { get; set; }
            public int CreatedSolidCount { get; set; }
            public bool MaterialApplied { get; set; }
            public string ErrorMessage { get; set; }
        }

        /// <summary>
        /// Phase A: .rfa を作成して保存する。プロジェクトのトランザクションを開く前に呼ぶこと。
        /// </summary>
        public static FamilyFileOutcome CreateFamilyFile(
            Autodesk.Revit.ApplicationServices.Application app,
            Document projectDoc,
            IList<Solid> solids,
            string templatePath,
            string savePath,
            ElementId materialId)
        {
            var outcome = new FamilyFileOutcome();
            Document famDoc = null;

            try
            {
                famDoc = app.NewFamilyDocument(templatePath);
                if (famDoc == null)
                {
                    outcome.ErrorMessage = "NewFamilyDocument returned null";
                    return outcome;
                }

                using (var t = new Transaction(famDoc, "Tools28 GenericModelMerge"))
                {
                    t.Start();

                    // テンプレートが一般モデル以外でも確実に一般モデルにする
                    try
                    {
                        var gm = Category.GetCategory(famDoc, BuiltInCategory.OST_GenericModel);
                        if (gm != null) famDoc.OwnerFamily.FamilyCategory = gm;
                    }
                    catch { }

                    // 材質はプロジェクトからファミリ文書へコピーしてから割り当てる
                    ElementId famMaterialId = ElementId.InvalidElementId;
                    if (materialId != null && materialId != ElementId.InvalidElementId)
                        famMaterialId = CopyMaterial(projectDoc, famDoc, materialId);

                    foreach (var solid in solids)
                    {
                        try
                        {
                            var ffe = FreeFormElement.Create(famDoc, solid);
                            if (ffe == null) { outcome.FailedSolidCount++; continue; }
                            outcome.CreatedSolidCount++;

                            if (famMaterialId != ElementId.InvalidElementId)
                            {
                                var p = ffe.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM);
                                if (p != null && !p.IsReadOnly && p.Set(famMaterialId))
                                    outcome.MaterialApplied = true;
                            }
                        }
                        catch
                        {
                            outcome.FailedSolidCount++;
                        }
                    }

                    t.Commit();
                }

                if (outcome.CreatedSolidCount == 0)
                {
                    outcome.ErrorMessage = "FreeFormElement could not be created";
                    return outcome;
                }

                var saveOpts = new SaveAsOptions { OverwriteExistingFile = true };
                famDoc.SaveAs(savePath, saveOpts);
                outcome.SavedPath = savePath;
                outcome.Success = true;
                return outcome;
            }
            catch (Exception ex)
            {
                outcome.ErrorMessage = ex.Message;
                return outcome;
            }
            finally
            {
                if (famDoc != null)
                {
                    try { famDoc.Close(false); } catch { }
                }
            }
        }

        /// <summary>
        /// Phase B: .rfa をロードして原点に配置する。プロジェクトのトランザクション内で呼ぶこと。
        /// </summary>
        public static ElementId LoadAndPlace(Document doc, string familyPath)
        {
            Family family;
            if (!doc.LoadFamily(familyPath, new OverwriteFamilyLoadOptions(), out family) || family == null)
            {
                // すでに同名でロード済みの場合は LoadFamily が false を返すことがある
                string name = Path.GetFileNameWithoutExtension(familyPath);
                family = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(f => f.Name == name);
                if (family == null) return ElementId.InvalidElementId;
            }

            var symbolId = family.GetFamilySymbolIds().FirstOrDefault();
            if (symbolId == null || symbolId == ElementId.InvalidElementId)
                return ElementId.InvalidElementId;

            var symbol = doc.GetElement(symbolId) as FamilySymbol;
            if (symbol == null) return ElementId.InvalidElementId;
            if (!symbol.IsActive) symbol.Activate();

            var instance = doc.Create.NewFamilyInstance(
                XYZ.Zero, symbol, StructuralType.NonStructural);
            return instance?.Id ?? ElementId.InvalidElementId;
        }

        private static ElementId CopyMaterial(Document source, Document dest, ElementId materialId)
        {
            try
            {
                var copied = ElementTransformUtils.CopyElements(
                    source, new List<ElementId> { materialId }, dest, Transform.Identity, null);
                return copied?.FirstOrDefault() ?? ElementId.InvalidElementId;
            }
            catch
            {
                return ElementId.InvalidElementId;
            }
        }

        /// <summary>
        /// 一般モデルのファミリテンプレート (.rft) を探す。
        /// Revit の言語・地域によってファイル名が違うため、既知の名前を順に探し、
        /// 見つからなければ null を返す（呼び出し側でユーザーに選ばせる）。
        /// </summary>
        public static string FindGenericModelTemplate(
            Autodesk.Revit.ApplicationServices.Application app)
        {
            string root = null;
            try { root = app.FamilyTemplatePath; } catch { }
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return null;

            string[] files;
            try { files = Directory.GetFiles(root, "*.rft", SearchOption.AllDirectories); }
            catch { return null; }
            if (files.Length == 0) return null;

            // 完全一致を優先し、次に部分一致で探す（メートル法のものを優先）
            string[] exact =
            {
                "一般モデル", "Metric Generic Model", "Generic Model",
                "公制常规模型", "常规模型",
            };

            foreach (var name in exact)
            {
                var hit = files.FirstOrDefault(f =>
                    string.Equals(Path.GetFileNameWithoutExtension(f), name,
                        StringComparison.OrdinalIgnoreCase));
                if (hit != null) return hit;
            }

            foreach (var name in exact)
            {
                var hit = files.FirstOrDefault(f =>
                    Path.GetFileNameWithoutExtension(f)
                        .IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);
                if (hit != null) return hit;
            }

            return null;
        }

        /// <summary>同名ファミリがロード済みでも上書きする。</summary>
        private class OverwriteFamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(
                Family sharedFamily, bool familyInUse,
                out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
        }
    }
}
