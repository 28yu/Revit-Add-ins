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
        /// <summary>設定が実際に反映されたビュー／ビューテンプレート数</summary>
        public int ViewCount { get; set; }

        /// <summary>反映したレイヤ設定の延べ件数</summary>
        public int LayerCount { get; set; }

        /// <summary>ビューへ直接書けず、テンプレートへ振り替えたビュー数</summary>
        public int TemplateFallbackViews { get; set; }

        /// <summary>実際に書き込んだビューテンプレート数</summary>
        public int TemplateFallbackCount { get; set; }

        /// <summary>移行元で既定から変更されていた設定の件数（0 なら移行しても見た目は変わらない）</summary>
        public int SourceSettingCount { get; set; }

        /// <summary>移行先の DWG に見つからなかったレイヤ名</summary>
        public List<string> MissingLayers { get; } = new List<string>();

        /// <summary>移行先に存在しなかった線種名</summary>
        public List<string> MissingLinePatterns { get; } = new List<string>();

        /// <summary>反映できなかったビューとその理由</summary>
        public List<string> Failures { get; } = new List<string>();
    }

    /// <summary>
    /// 読み取った DWG レイヤ表示設定を、移行先モデルのビュー／ビューテンプレートへ書き込むサービス。
    ///
    /// ElementId はモデル間で通用しないため、
    ///   - DWG      : ダイアログで選んだ移行先 DWG
    ///   - レイヤ    : サブカテゴリ名の一致
    ///   - 線種      : LinePatternElement の名前一致
    /// で移行先の ID へ解決し直す。解決できなかったものはレポートに残す。
    ///
    /// ⚠️ 書き込みは「書けたか」ではなく「効いたか」で判定する。
    /// ビューテンプレートに制御されているビューへの SetCategoryOverrides は
    /// 例外を投げずに黙って無視されることがあり、書きっぱなしでは成功と区別できない。
    /// そのため 1 カテゴリごとに読み戻して検証し、効いていなければ
    /// そのビューのビューテンプレートへ書き込み先を振り替える。
    /// </summary>
    public class DwgLayerApplier
    {
        private Dictionary<string, ElementId> _linePatterns;

        /// <summary>1要素（ビュー／テンプレート）への書き込み結果。</summary>
        private sealed class WriteStats
        {
            /// <summary>書き込んで実際に効いた件数</summary>
            public int Applied;
            /// <summary>書き込んだが値が変わらなかった件数（テンプレート制御下など）</summary>
            public int Ineffective;
            /// <summary>例外で書き込めなかった件数</summary>
            public int Failed;
            /// <summary>最初に起きたエラーの内容（診断用）</summary>
            public string FirstError;

            public int Attempted => Applied + Ineffective + Failed;
        }

        /// <summary>
        /// 移行元 1 DWG 分のレイヤ設定を、移行先の 1 DWG × 複数ビューへ適用する。
        /// トランザクションは内部で開始・コミットする。
        /// </summary>
        /// <param name="targetDoc">移行先ドキュメント</param>
        /// <param name="sourceLayers">移行元のレイヤ名 -&gt; 設定（"" は DWG 本体）</param>
        /// <param name="targetViews">
        /// 適用先のビュー。チェックされたものをそのまま渡してよい（重複排除は不要）。
        /// 直接書けないビューは、このメソッドがビューテンプレートへ振り替える。
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
            result.SourceSettingCount = DwgLayerScanner.CountConfigured(sourceLayers.Values);

            DiagLog.Write($"[DwgVg] 適用開始 ビュー={targetViews.Count} レイヤ設定={sourceLayers.Count} " +
                          $"(既定以外={result.SourceSettingCount}) 移行先DWG='{targetDwg.Name}' " +
                          $"レイヤ数={targetDwg.Layers.Count} 線種={_linePatterns.Count}");

            var missingLayers = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            var missingPatterns = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            // 同じテンプレートを共有するビューが複数あっても、書き込みは1回で済ませる
            var writtenTemplates = new HashSet<ElementId>();

            using (var t = new Transaction(targetDoc, Loc.S("DwgVg.Txn.Apply")))
            {
                t.Start();

                foreach (var entry in targetViews)
                {
                    if (entry == null) continue;

                    if (!(targetDoc.GetElement(entry.Id) is View view))
                    {
                        result.Failures.Add(string.Format(Loc.S("DwgVg.Fail.ViewNotFound"), entry.Name));
                        continue;
                    }

                    var w = WriteAll(view, sourceLayers, targetDwg, missingLayers, missingPatterns);
                    DiagLog.Write($"[DwgVg]   '{entry.Name}' 直接: 反映={w.Applied} 未反映={w.Ineffective} " +
                                  $"失敗={w.Failed} {w.FirstError}");

                    if (w.Applied > 0)
                    {
                        result.ViewCount++;
                        result.LayerCount += w.Applied;
                        continue;
                    }

                    if (w.Attempted == 0)
                    {
                        // 移行先 DWG に一致するレイヤが1つも無い（MissingLayers に記録済み）
                        result.Failures.Add(string.Format(Loc.S("DwgVg.Fail.NoLayerMatch"), entry.Name));
                        continue;
                    }

                    // ビュー側では効かなかった。テンプレートに制御されているとみなして振り替える
                    ApplyViaTemplate(targetDoc, entry, view, sourceLayers, targetDwg,
                                     missingLayers, missingPatterns, writtenTemplates, w, result);
                }

                t.Commit();
            }

            result.MissingLayers.AddRange(missingLayers.OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase));
            result.MissingLinePatterns.AddRange(missingPatterns.OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase));

            DiagLog.Write($"[DwgVg] 適用終了 反映ビュー={result.ViewCount} 反映レイヤ={result.LayerCount} " +
                          $"テンプレート振替={result.TemplateFallbackViews}ビュー/{result.TemplateFallbackCount}件 " +
                          $"未一致レイヤ={result.MissingLayers.Count} 失敗={result.Failures.Count}");
            return result;
        }

        /// <summary>
        /// ビューへ直接書けなかった場合に、そのビューのビューテンプレートへ書き込む。
        /// テンプレートが無い／テンプレートでも効かない場合は失敗として記録する。
        /// </summary>
        private void ApplyViaTemplate(
            Document targetDoc, ViewEntry entry, View view,
            IDictionary<string, LayerGraphicSetting> sourceLayers, DwgDefinition targetDwg,
            HashSet<string> missingLayers, HashSet<string> missingPatterns,
            HashSet<ElementId> writtenTemplates, WriteStats direct, TransferResult result)
        {
            // ビューテンプレート自身に書いて効かなかった場合は、振り替え先が無い
            if (entry.IsTemplate)
            {
                result.Failures.Add(string.Format(
                    Loc.S("DwgVg.Fail.NotApplied"), entry.Name, direct.FirstError ?? ""));
                return;
            }

            var tpl = DwgLayerScanner.GetAssignedTemplate(targetDoc, view);
            if (tpl == null)
            {
                result.Failures.Add(string.Format(
                    Loc.S("DwgVg.Fail.NotApplied"), entry.Name, direct.FirstError ?? ""));
                return;
            }

            result.TemplateFallbackViews++;

            // 同じテンプレートには一度だけ書けばよい（2件目以降も反映済みとして扱う）
            if (!writtenTemplates.Add(tpl.Id)) return;

            string tplName;
            try { tplName = tpl.Name ?? ""; } catch { tplName = ""; }

            var wt = WriteAll(tpl, sourceLayers, targetDwg, missingLayers, missingPatterns);
            DiagLog.Write($"[DwgVg]   → テンプレート '{tplName}' へ振替: 反映={wt.Applied} " +
                          $"未反映={wt.Ineffective} 失敗={wt.Failed} {wt.FirstError}");

            if (wt.Applied > 0)
            {
                result.TemplateFallbackCount++;
                result.ViewCount++;
                result.LayerCount += wt.Applied;
            }
            else
            {
                result.Failures.Add(string.Format(
                    Loc.S("DwgVg.Fail.TemplateNotApplied"), entry.Name, tplName, wt.FirstError ?? ""));
            }
        }

        /// <summary>1要素へ全レイヤ分の設定を書き込む。</summary>
        private WriteStats WriteAll(
            View view, IDictionary<string, LayerGraphicSetting> sourceLayers, DwgDefinition targetDwg,
            HashSet<string> missingLayers, HashSet<string> missingPatterns)
        {
            var st = new WriteStats();

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

                WriteOne(view, targetCatId, setting, missingPatterns, st);
            }

            return st;
        }

        /// <summary>1カテゴリ分の設定を書き込み、実際に効いたかを読み戻して確認する。</summary>
        private void WriteOne(
            View view, ElementId categoryId, LayerGraphicSetting s,
            HashSet<string> missingPatterns, WriteStats st)
        {
            if (categoryId == null || categoryId == ElementId.InvalidElementId) return;

            try
            {
                view.SetCategoryOverrides(categoryId, BuildOverrides(s, missingPatterns));
            }
            catch (Exception ex)
            {
                st.Failed++;
                if (st.FirstError == null) st.FirstError = ex.Message;
                return;
            }

            // 表示/非表示は別 API
            try
            {
                if (view.CanCategoryBeHidden(categoryId) && view.GetCategoryHidden(categoryId) != s.Hidden)
                    view.SetCategoryHidden(categoryId, s.Hidden);
            }
            catch (Exception ex)
            {
                if (st.FirstError == null) st.FirstError = ex.Message;
            }

            if (IsEffective(view, categoryId, s)) st.Applied++;
            else st.Ineffective++;
        }

        /// <summary>移行元の設定から、書き込む OverrideGraphicSettings を組み立てる。</summary>
        private OverrideGraphicSettings BuildOverrides(LayerGraphicSetting s, HashSet<string> missingPatterns)
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
            return ogs;
        }

        /// <summary>
        /// 書き込んだ設定がそのビューで実際に効いているかを判定する。
        ///
        /// ビューテンプレートに制御されているビューでは GetCategoryOverrides が
        /// テンプレート側の値を返すため、書いた値と食い違えば「効いていない」と分かる。
        /// 線種は移行先に同名が無いと書き込まない（別途レポートする）ため比較対象から外す。
        /// </summary>
        private static bool IsEffective(View view, ElementId categoryId, LayerGraphicSetting s)
        {
            try
            {
                if (view.CanCategoryBeHidden(categoryId) && view.GetCategoryHidden(categoryId) != s.Hidden)
                    return false;

                var cur = view.GetCategoryOverrides(categoryId);
                if (cur == null) return false;

                if (cur.Halftone != s.Halftone) return false;

                if (!SameColor(cur.ProjectionLineColor, s.ProjectionLineColor)) return false;
                if (!SameColor(cur.CutLineColor, s.CutLineColor)) return false;

                if (cur.ProjectionLineWeight != (s.ProjectionLineWeight > 0 ? s.ProjectionLineWeight : -1))
                    return false;
                if (cur.CutLineWeight != (s.CutLineWeight > 0 ? s.CutLineWeight : -1))
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>色の一致判定。どちらも未設定なら一致とみなす。</summary>
        private static bool SameColor(Color current, Color wanted)
        {
            bool curOk = current != null && current.IsValid;
            bool wantOk = wanted != null && wanted.IsValid;
            if (curOk != wantOk) return false;
            if (!curOk) return true;

            return current.Red == wanted.Red
                && current.Green == wanted.Green
                && current.Blue == wanted.Blue;
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
