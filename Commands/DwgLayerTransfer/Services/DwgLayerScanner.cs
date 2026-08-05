using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.DwgLayerTransfer.Models;
using Tools28.Localization;

namespace Tools28.Commands.DwgLayerTransfer.Services
{
    /// <summary>
    /// モデル内の DWG（読み込み／リンク）とそのレイヤを列挙し、
    /// 各ビュー／ビューテンプレートに設定されている表示設定を読み取るサービス。
    ///
    /// 走査方針:
    ///   - DWG カテゴリは Document.Settings.Categories のうち「親を持たず、ID が正」のものを採る。
    ///     組み込みカテゴリの ID は負のため、これだけで読み込み DWG のカテゴリを抽出できる。
    ///   - リンク／読み込みの別は ImportInstance.IsLinked から補足する（表示用のみ）。
    ///   - 設定値は View.GetCategoryHidden / GetCategoryOverrides で読む。これは V/G ダイアログの
    ///     「読み込みカテゴリ」タブが書き込む先と同一。
    /// </summary>
    public class DwgLayerScanner
    {
        // ===== DWG カテゴリの列挙 =====

        /// <summary>モデル内の DWG（カテゴリ）とレイヤ（サブカテゴリ）を列挙する。</summary>
        public List<DwgDefinition> EnumerateDwgs(Document doc)
        {
            var result = new List<DwgDefinition>();
            if (doc == null) return result;

            // リンク/読み込みの別を先に集める（表示用の補足情報）
            var linkedFlags = CollectImportKinds(doc);

            Categories cats;
            try { cats = doc.Settings.Categories; }
            catch { return result; }

            foreach (Category cat in cats)
            {
                if (cat == null) continue;

                // 組み込みカテゴリ（ID が負）とサブカテゴリ（親あり）を除外し、
                // 読み込み DWG のカテゴリだけを残す
                int idValue;
                try { idValue = cat.Id.IntValue(); }
                catch { continue; }
                if (idValue <= 0) continue;

                try { if (cat.Parent != null) continue; }
                catch { continue; }

                string name;
                try { name = cat.Name; }
                catch { continue; }
                if (string.IsNullOrEmpty(name)) continue;

                var def = new DwgDefinition { Name = name, CategoryId = cat.Id };
                if (linkedFlags.TryGetValue(idValue, out bool isLinked))
                    def.IsLinked = isLinked;

                try
                {
                    foreach (Category sub in cat.SubCategories)
                    {
                        if (sub == null) continue;
                        string sname;
                        try { sname = sub.Name; } catch { continue; }
                        if (string.IsNullOrEmpty(sname)) continue;
                        def.Layers[sname] = sub.Id;
                    }
                }
                catch { }

                result.Add(def);
            }

            return result
                .OrderBy(d => d.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        /// <summary>ImportInstance を走査して カテゴリID -&gt; リンクかどうか のマップを作る。</summary>
        private static Dictionary<int, bool> CollectImportKinds(Document doc)
        {
            var map = new Dictionary<int, bool>();
            try
            {
                var instances = new FilteredElementCollector(doc)
                    .OfClass(typeof(ImportInstance))
                    .Cast<ImportInstance>();

                foreach (var inst in instances)
                {
                    if (inst == null) continue;
                    Category cat;
                    try { cat = inst.Category; } catch { continue; }
                    if (cat == null) continue;

                    try { map[cat.Id.IntValue()] = inst.IsLinked; }
                    catch { }
                }
            }
            catch { }
            return map;
        }

        // ===== ビューの列挙 =====

        /// <summary>
        /// ビューまたはビューテンプレートを列挙する。
        /// </summary>
        /// <param name="templates">true ならビューテンプレートのみ、false なら通常ビューのみ</param>
        /// <param name="checkTemplateControl">
        /// true の場合、ビューテンプレートが「読み込みカテゴリ」の V/G を制御しているビューに
        /// 適用不可の理由を設定する（移行先の列挙時に使う）。
        /// </param>
        public List<ViewEntry> EnumerateViews(Document doc, bool templates, bool checkTemplateControl)
        {
            var result = new List<ViewEntry>();
            if (doc == null) return result;

            IEnumerable<View> views;
            try
            {
                views = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>();
            }
            catch { return result; }

            foreach (var v in views)
            {
                if (v == null) continue;

                bool isTemplate;
                try { isTemplate = v.IsTemplate; } catch { continue; }
                if (isTemplate != templates) continue;

                if (isTemplate)
                {
                    // ビューテンプレートに AreGraphicsOverridesAllowed() は使えないため型で除外する
                    if (v is ViewSchedule || v is ViewSheet) continue;
                }
                else
                {
                    // グラフィックス上書きに対応しないビュー（集計表・シート等）は対象外
                    try { if (!v.AreGraphicsOverridesAllowed()) continue; }
                    catch { continue; }
                }

                string name;
                try { name = v.Name; } catch { continue; }
                if (string.IsNullOrEmpty(name)) continue;

                var entry = new ViewEntry { Id = v.Id, Name = name, IsTemplate = isTemplate };

                if (checkTemplateControl)
                {
                    if (isTemplate)
                    {
                        // 「読み込みカテゴリ」を制御していないテンプレートに書き込んでも
                        // ビューには反映されないため、はっきり伝える
                        if (!TemplateControlsImportVg(v))
                            entry.BlockReason = Loc.S("DwgVg.Status.TemplateNotControlling");
                    }
                    else if (IsImportVgControlledByTemplate(doc, v))
                    {
                        entry.BlockReason = Loc.S("DwgVg.Status.TemplateControlled");
                    }
                }

                result.Add(entry);
            }

            return result
                .OrderBy(e => e.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// ビューに割り当てられたビューテンプレートが「読み込みカテゴリの V/G」を制御しているか。
        /// 制御されている場合、ビュー側へ SetCategoryOverrides しても反映されない。
        /// </summary>
        public static bool IsImportVgControlledByTemplate(Document doc, View v)
        {
            try
            {
                ElementId tid = v.ViewTemplateId;
                if (tid == null || tid == ElementId.InvalidElementId) return false;

                if (!(doc.GetElement(tid) is View tpl)) return false;
                return TemplateControlsImportVg(tpl);
            }
            catch
            {
                // 判定できない場合は制御されていない扱いにし、適用時の例外で個別に処理する
                return false;
            }
        }

        /// <summary>
        /// ビューテンプレートが「読み込みカテゴリの V/G」を制御しているか。
        /// GetNonControlledTemplateParameterIds は「制御していない」項目を返すため、
        /// そこに含まれていなければ制御下と判断する。
        /// </summary>
        public static bool TemplateControlsImportVg(View tpl)
        {
            try
            {
                var nonControlled = tpl.GetNonControlledTemplateParameterIds();
                if (nonControlled == null) return true;

                var importVg = new ElementId(BuiltInParameter.VIS_GRAPHICS_IMPORT);
                return !nonControlled.Any(id => id == importVg);
            }
            catch
            {
                // 判定できない場合は制御されている前提にして、余計な警告を出さない
                return true;
            }
        }

        // ===== 設定値の読み取り =====

        /// <summary>
        /// 指定ビュー群について、全 DWG レイヤの表示設定スナップショットを作る。
        /// </summary>
        public List<ViewSettingSnapshot> ReadSettings(
            Document doc, IEnumerable<ViewEntry> views, IEnumerable<DwgDefinition> dwgs)
        {
            var snapshots = new List<ViewSettingSnapshot>();
            if (doc == null || views == null || dwgs == null) return snapshots;

            var dwgList = dwgs.ToList();

            foreach (var entry in views)
            {
                if (!(doc.GetElement(entry.Id) is View view)) continue;

                var snap = new ViewSettingSnapshot { View = entry };

                foreach (var dwg in dwgList)
                {
                    var layers = new Dictionary<string, LayerGraphicSetting>(
                        StringComparer.CurrentCultureIgnoreCase);

                    // DWG 本体（親カテゴリ）はレイヤ名 "" で保持する
                    var root = ReadOne(doc, view, dwg.CategoryId, "");
                    if (root != null) layers[""] = root;

                    foreach (var kv in dwg.Layers)
                    {
                        var s = ReadOne(doc, view, kv.Value, kv.Key);
                        if (s != null) layers[kv.Key] = s;
                    }

                    if (layers.Count > 0) snap.ByDwg[dwg.Name] = layers;
                }

                snapshots.Add(snap);
            }

            return snapshots;
        }

        /// <summary>1カテゴリ（DWG本体またはレイヤ）分の設定を読む。読めない場合は null。</summary>
        private static LayerGraphicSetting ReadOne(
            Document doc, View view, ElementId categoryId, string layerName)
        {
            if (categoryId == null || categoryId == ElementId.InvalidElementId) return null;

            var s = new LayerGraphicSetting { LayerName = layerName };

            try { s.Hidden = view.GetCategoryHidden(categoryId); }
            catch { /* 非表示にできないカテゴリは既定（表示）扱い */ }

            OverrideGraphicSettings ogs;
            try { ogs = view.GetCategoryOverrides(categoryId); }
            catch { return s; }
            if (ogs == null) return s;

            try
            {
                s.ProjectionLineColor = ogs.ProjectionLineColor;
                s.ProjectionLineWeight = ogs.ProjectionLineWeight;
                s.ProjectionLinePattern = LinePatternName(doc, ogs.ProjectionLinePatternId);

                s.CutLineColor = ogs.CutLineColor;
                s.CutLineWeight = ogs.CutLineWeight;
                s.CutLinePattern = LinePatternName(doc, ogs.CutLinePatternId);

                s.Halftone = ogs.Halftone;
            }
            catch { }

            return s;
        }

        /// <summary>
        /// 線種 ElementId を移行先で解決できる「名前」に変換する。
        /// 上書きなしは null、既定の実線は予約名を返す。
        /// </summary>
        private static string LinePatternName(Document doc, ElementId id)
        {
            if (id == null || id == ElementId.InvalidElementId) return null;

            try
            {
                if (id == LinePatternElement.GetSolidPatternId())
                    return LayerGraphicSetting.SolidPatternMarker;
            }
            catch { }

            try
            {
                if (doc.GetElement(id) is LinePatternElement lpe && !string.IsNullOrEmpty(lpe.Name))
                    return lpe.Name;
            }
            catch { }

            return null;
        }
    }
}
