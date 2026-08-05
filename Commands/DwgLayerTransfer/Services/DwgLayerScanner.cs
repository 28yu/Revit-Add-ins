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

            // ImportInstance の配置情報（リンク/読み込みの別・どのビューにあるか）を先に集める
            var placements = CollectImportPlacements(doc);

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
                if (placements.TryGetValue(idValue, out var placement))
                {
                    def.IsLinked = placement.IsLinked;
                    def.AppearsEverywhere = placement.HasModelWide;
                    foreach (var vid in placement.OwnerViews) def.OwnerViewIds.Add(vid);
                }

                // ImportInstance が1つも見つからなかった DWG は、どのビューでも選べるようにしておく
                // （インスタンスを消してもカテゴリは残るため、取りこぼすと設定を移せなくなる）
                if (!def.AppearsEverywhere && def.OwnerViewIds.Count == 0)
                    def.AppearsEverywhere = true;

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

        /// <summary>ImportInstance 1カテゴリ分の配置情報。</summary>
        private sealed class ImportPlacement
        {
            public bool? IsLinked;
            /// <summary>モデル全体に配置されたインスタンスがあるか（OwnerViewId が無効）</summary>
            public bool HasModelWide;
            /// <summary>ビュー固有に貼り付けられたインスタンスの所有ビュー</summary>
            public readonly HashSet<ElementId> OwnerViews = new HashSet<ElementId>();
        }

        /// <summary>
        /// ImportInstance を走査して カテゴリID -&gt; 配置情報 のマップを作る。
        ///
        /// OwnerViewId が無効なインスタンスは「モデルに読み込み／リンクされた DWG」で全ビューに現れうる。
        /// 有効なインスタンスは「そのビューにだけ貼り付けた DWG」で、当該ビューにしか現れない。
        /// </summary>
        private static Dictionary<int, ImportPlacement> CollectImportPlacements(Document doc)
        {
            var map = new Dictionary<int, ImportPlacement>();
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

                    int key;
                    try { key = cat.Id.IntValue(); }
                    catch { continue; }

                    if (!map.TryGetValue(key, out var placement))
                    {
                        placement = new ImportPlacement();
                        map[key] = placement;
                    }

                    try { placement.IsLinked = inst.IsLinked; } catch { }

                    try
                    {
                        ElementId owner = inst.OwnerViewId;
                        if (owner == null || owner == ElementId.InvalidElementId)
                            placement.HasModelWide = true;
                        else
                            placement.OwnerViews.Add(owner);
                    }
                    catch
                    {
                        // 判定できない場合は全ビューに現れる扱いにする（取りこぼし防止）
                        placement.HasModelWide = true;
                    }
                }
            }
            catch { }
            return map;
        }

        /// <summary>
        /// 指定ビューに現れる DWG だけに絞り込む。
        /// ビューテンプレートはジオメトリを持たず全 DWG のカテゴリを制御するため、絞り込まない。
        /// </summary>
        public static List<DwgDefinition> FilterForView(
            IEnumerable<DwgDefinition> dwgs, ElementId viewId, bool isTemplate)
        {
            if (dwgs == null) return new List<DwgDefinition>();
            if (isTemplate) return dwgs.ToList();
            return dwgs.Where(d => d.AppearsInView(viewId)).ToList();
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
                    else
                    {
                        // テンプレートに制御されているビューは、ビュー側から直接変更できない。
                        // 使えないものとして弾かず、書き込み先をそのテンプレートへ振り替える
                        var tpl = GetControllingTemplate(doc, v);
                        if (tpl != null)
                        {
                            string tplName;
                            try { tplName = tpl.Name ?? ""; } catch { tplName = ""; }

                            entry.TemplateId = tpl.Id;
                            entry.TemplateName = tplName;
                            entry.Note = string.Format(Loc.S("DwgVg.Status.ViaTemplate"), tplName);
                        }
                    }
                }

                result.Add(entry);
            }

            return result
                .OrderBy(e => e.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// ビューの「読み込みカテゴリの V/G」を制御しているビューテンプレートを返す（無ければ null）。
        /// 制御されている場合、ビュー側へ SetCategoryOverrides しても反映されないため、
        /// 呼び出し側は書き込み先をこのテンプレートへ振り替える。
        /// </summary>
        public static View GetControllingTemplate(Document doc, View v)
        {
            var tpl = GetAssignedTemplate(doc, v);
            if (tpl == null) return null;

            try { return TemplateControlsImportVg(tpl) ? tpl : null; }
            catch { return null; }
        }

        /// <summary>
        /// ビューに割り当てられているビューテンプレートを返す（無ければ null）。
        /// 「読み込みカテゴリを制御しているか」は問わない。
        /// 適用時にビューへ直接書けなかった場合の振り替え先を探すのに使う。
        /// </summary>
        public static View GetAssignedTemplate(Document doc, View v)
        {
            try
            {
                if (doc == null || v == null) return null;

                ElementId tid = v.ViewTemplateId;
                if (tid == null || tid == ElementId.InvalidElementId) return null;

                return doc.GetElement(tid) as View;
            }
            catch
            {
                return null;
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
        /// 1ビュー・1 DWG 分のレイヤ表示設定を読み取る。
        /// 戻り値のキーはレイヤ名で、"" は DWG 本体（親カテゴリ）を表す。
        /// </summary>
        public Dictionary<string, LayerGraphicSetting> ReadSettings(
            Document doc, ElementId viewId, DwgDefinition dwg)
        {
            var layers = new Dictionary<string, LayerGraphicSetting>(
                StringComparer.CurrentCultureIgnoreCase);

            if (doc == null || dwg == null) return layers;
            if (!(doc.GetElement(viewId) is View view)) return layers;

            var root = ReadOne(doc, view, dwg.CategoryId, "");
            if (root != null) layers[""] = root;

            foreach (var kv in dwg.Layers)
            {
                var s = ReadOne(doc, view, kv.Value, kv.Key);
                if (s != null) layers[kv.Key] = s;
            }

            return layers;
        }

        /// <summary>既定から変化している設定の件数を数える（一覧の「設定数」表示用）。</summary>
        public static int CountConfigured(IEnumerable<LayerGraphicSetting> settings)
            => settings?.Count(s => s != null && s.HasAnySetting) ?? 0;

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

                s.SurfaceFgPattern = FillPatternName(doc, ogs.SurfaceForegroundPatternId);
                s.SurfaceFgColor = ogs.SurfaceForegroundPatternColor;
                s.SurfaceFgVisible = ogs.IsSurfaceForegroundPatternVisible;

                s.SurfaceBgPattern = FillPatternName(doc, ogs.SurfaceBackgroundPatternId);
                s.SurfaceBgColor = ogs.SurfaceBackgroundPatternColor;
                s.SurfaceBgVisible = ogs.IsSurfaceBackgroundPatternVisible;

                s.CutFgPattern = FillPatternName(doc, ogs.CutForegroundPatternId);
                s.CutFgColor = ogs.CutForegroundPatternColor;
                s.CutFgVisible = ogs.IsCutForegroundPatternVisible;

                s.CutBgPattern = FillPatternName(doc, ogs.CutBackgroundPatternId);
                s.CutBgColor = ogs.CutBackgroundPatternColor;
                s.CutBgVisible = ogs.IsCutBackgroundPatternVisible;

                s.Transparency = ogs.Transparency;
                s.Halftone = ogs.Halftone;
                s.DetailLevel = ogs.DetailLevel;
            }
            catch { }

            return s;
        }

        /// <summary>塗潰しパターン ElementId を移行先で解決できる「名前」に変換する。</summary>
        private static string FillPatternName(Document doc, ElementId id)
        {
            try
            {
                if (id == null || id == ElementId.InvalidElementId) return null;
                return (doc?.GetElement(id) as FillPatternElement)?.Name;
            }
            catch { return null; }
        }

        /// <summary>
        /// 1カテゴリ分の設定を、Revit が持っている項目すべて 1 行の文字列にする。
        /// 移行元と移行先を突き合わせて「まだ何が違うのか」を目視で確かめるための診断用。
        /// </summary>
        public static string DescribeOverride(Document doc, ElementId viewId, ElementId categoryId)
        {
            try
            {
                if (!(doc.GetElement(viewId) is View view)) return "ビュー取得不可";

                string hidden;
                try { hidden = view.GetCategoryHidden(categoryId).ToString(); }
                catch (Exception ex) { hidden = "?(" + ex.Message + ")"; }

                var o = view.GetCategoryOverrides(categoryId);
                if (o == null) return $"非表示={hidden} 上書き=なし";

                string C(Color c) => (c != null && c.IsValid) ? $"{c.Red},{c.Green},{c.Blue}" : "-";
                string F(ElementId id) => FillPatternName(doc, id) ?? "-";
                string L(ElementId id) => LinePatternName(doc, id) ?? "-";

                return $"非表示={hidden} 投影線(色={C(o.ProjectionLineColor)} 幅={o.ProjectionLineWeight} " +
                       $"種={L(o.ProjectionLinePatternId)}) 切断線(色={C(o.CutLineColor)} 幅={o.CutLineWeight} " +
                       $"種={L(o.CutLinePatternId)}) 面前景({F(o.SurfaceForegroundPatternId)} 色={C(o.SurfaceForegroundPatternColor)} " +
                       $"表示={o.IsSurfaceForegroundPatternVisible}) 面背景({F(o.SurfaceBackgroundPatternId)} " +
                       $"色={C(o.SurfaceBackgroundPatternColor)} 表示={o.IsSurfaceBackgroundPatternVisible}) " +
                       $"断前景({F(o.CutForegroundPatternId)} 色={C(o.CutForegroundPatternColor)} 表示={o.IsCutForegroundPatternVisible}) " +
                       $"断背景({F(o.CutBackgroundPatternId)} 色={C(o.CutBackgroundPatternColor)} 表示={o.IsCutBackgroundPatternVisible}) " +
                       $"透過={o.Transparency} HT={o.Halftone} 詳細={o.DetailLevel}";
            }
            catch (Exception ex)
            {
                return "取得失敗: " + ex.Message;
            }
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
