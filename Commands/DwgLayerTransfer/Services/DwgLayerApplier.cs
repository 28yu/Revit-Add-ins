using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Tools28.Commands.DwgLayerTransfer.Models;
using Tools28.Localization;

namespace Tools28.Commands.DwgLayerTransfer.Services
{
    /// <summary>移行結果のレポート。</summary>
    public sealed class TransferResult
    {
        /// <summary>設定を書き込んだビュー／ビューテンプレート数</summary>
        public int ViewCount { get; set; }

        /// <summary>書き込んだレイヤ設定の延べ件数</summary>
        public int LayerCount { get; set; }

        /// <summary>移行先の DWG に見つからなかったレイヤ名</summary>
        public List<string> MissingLayers { get; } = new List<string>();

        /// <summary>移行先に存在しなかった線種名</summary>
        public List<string> MissingLinePatterns { get; } = new List<string>();
    }

    /// <summary>
    /// 読み取った DWG レイヤ表示設定を、移行先モデルのビュー／ビューテンプレートへ書き込むサービス。
    ///
    /// ElementId はモデル間で通用しないため、
    ///   - DWG      : ダイアログで選んだ移行先 DWG
    ///   - レイヤ    : サブカテゴリ名の一致
    ///   - 線種      : LinePatternElement の名前一致
    /// で移行先の ID へ解決し直す。解決できなかったものはレポートに残す。
    /// </summary>
    public class DwgLayerApplier
    {
        private Dictionary<string, ElementId> _linePatterns;

        /// <summary>
        /// 移行元 1 DWG 分のレイヤ設定を、移行先の 1 DWG × 複数ビューへ適用する。
        /// トランザクションは内部で開始・コミットする。
        /// </summary>
        /// <param name="targetDoc">移行先ドキュメント</param>
        /// <param name="sourceLayers">移行元のレイヤ名 -&gt; 設定（"" は DWG 本体）</param>
        /// <param name="targetViews">
        /// 適用先のビュー。適用可能なものだけを、書き込み先（<see cref="ViewEntry.WriteTargetId"/>）で
        /// 重複排除した状態で渡すこと。テンプレート制御下のビューはそのテンプレートへ書き込まれる。
        /// </param>
        /// <param name="targetDwg">適用先の DWG</param>
        public TransferResult Apply(
            Document targetDoc,
            IDictionary<string, LayerGraphicSetting> sourceLayers,
            IList<ViewEntry> targetViews,
            DwgDefinition targetDwg)
        {
            var result = new TransferResult();
            if (targetDoc == null || sourceLayers == null || sourceLayers.Count == 0) return result;
            if (targetViews == null || targetViews.Count == 0 || targetDwg == null) return result;

            _linePatterns = BuildLinePatternMap(targetDoc);

            var missingLayers = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            var missingPatterns = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            using (var t = new Transaction(targetDoc, Loc.S("DwgVg.Txn.Apply")))
            {
                t.Start();

                foreach (var entry in targetViews)
                {
                    if (entry == null) continue;
                    // テンプレートに制御されているビューは、そのテンプレートが書き込み先になる
                    if (!(targetDoc.GetElement(entry.WriteTargetId) is View targetView)) continue;

                    int appliedInView = 0;

                    foreach (var kv in sourceLayers)
                    {
                        string layerName = kv.Key;
                        LayerGraphicSetting setting = kv.Value;
                        if (setting == null) continue;

                        // レイヤ名 "" は DWG 本体（親カテゴリ）
                        ElementId targetCatId;
                        if (layerName.Length == 0)
                        {
                            targetCatId = targetDwg.CategoryId;
                        }
                        else if (!targetDwg.Layers.TryGetValue(layerName, out targetCatId))
                        {
                            missingLayers.Add(layerName);
                            continue;
                        }

                        if (ApplyOne(targetView, targetCatId, setting, missingPatterns))
                            appliedInView++;
                    }

                    if (appliedInView > 0)
                    {
                        result.ViewCount++;
                        result.LayerCount += appliedInView;
                    }
                }

                t.Commit();
            }

            result.MissingLayers.AddRange(missingLayers.OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase));
            result.MissingLinePatterns.AddRange(missingPatterns.OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase));
            return result;
        }

        /// <summary>1カテゴリ分の設定を書き込む。何かしら書き込めたら true。</summary>
        private bool ApplyOne(
            View view, ElementId categoryId, LayerGraphicSetting s, HashSet<string> missingPatterns)
        {
            if (categoryId == null || categoryId == ElementId.InvalidElementId) return false;

            bool applied = false;

            // --- 表示/非表示 ---
            try
            {
                if (view.CanCategoryBeHidden(categoryId))
                {
                    view.SetCategoryHidden(categoryId, s.Hidden);
                    applied = true;
                }
            }
            catch { }

            // --- 上書き（色・線幅・線種・ハーフトーン）---
            try
            {
                var ogs = new OverrideGraphicSettings();

                if (s.ProjectionLineColor != null && s.ProjectionLineColor.IsValid)
                    ogs.SetProjectionLineColor(s.ProjectionLineColor);
                if (s.ProjectionLineWeight > 0)
                    ogs.SetProjectionLineWeight(s.ProjectionLineWeight);
                var projPattern = ResolveLinePattern(s.ProjectionLinePattern, missingPatterns);
                if (projPattern != null)
                    ogs.SetProjectionLinePatternId(projPattern);

                if (s.CutLineColor != null && s.CutLineColor.IsValid)
                    ogs.SetCutLineColor(s.CutLineColor);
                if (s.CutLineWeight > 0)
                    ogs.SetCutLineWeight(s.CutLineWeight);
                var cutPattern = ResolveLinePattern(s.CutLinePattern, missingPatterns);
                if (cutPattern != null)
                    ogs.SetCutLinePatternId(cutPattern);

                ogs.SetHalftone(s.Halftone);

                view.SetCategoryOverrides(categoryId, ogs);
                applied = true;
            }
            catch { }

            return applied;
        }

        /// <summary>
        /// 線種名を移行先の ElementId へ解決する。
        /// null（上書きなし）と未解決はいずれも null を返し、未解決のみレポートに記録する。
        /// </summary>
        private ElementId ResolveLinePattern(string name, HashSet<string> missingPatterns)
        {
            if (name == null) return null;                       // 上書きなし

            if (name == LayerGraphicSetting.SolidPatternMarker)
            {
                try { return LinePatternElement.GetSolidPatternId(); }
                catch { return null; }
            }

            if (_linePatterns.TryGetValue(name, out var id)) return id;

            missingPatterns.Add(name);
            return null;
        }

        private static Dictionary<string, ElementId> BuildLinePatternMap(Document doc)
        {
            var map = new Dictionary<string, ElementId>(StringComparer.CurrentCultureIgnoreCase);
            try
            {
                foreach (var e in new FilteredElementCollector(doc)
                             .OfClass(typeof(LinePatternElement))
                             .ToElements())
                {
                    if (e is LinePatternElement lpe && !string.IsNullOrEmpty(lpe.Name))
                        map[lpe.Name] = lpe.Id;
                }
            }
            catch { }
            return map;
        }
    }
}
