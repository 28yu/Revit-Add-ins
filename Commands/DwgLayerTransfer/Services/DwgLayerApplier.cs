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

        /// <summary>Revit が非表示化を許可せず、表示/非表示を再現できなかった件数</summary>
        public int HiddenBlocked { get; set; }

        /// <summary>オブジェクトスタイルを書き換えたレイヤ数（モデル全体に効く）</summary>
        public int ObjectStyleCount { get; set; }

        /// <summary>反映できなかったビューとその理由</summary>
        public List<string> Failures { get; } = new List<string>();
    }

    /// <summary>
    /// 読み取った DWG レイヤ表示設定を、移行先モデルのビュー／ビューテンプレートへ書き込むサービス。
    ///
    /// ElementId はモデル間で通用しないため、
    ///   - DWG      : ダイアログで選んだ移行先 DWG（複数可）
    ///   - レイヤ    : サブカテゴリ名の一致
    ///   - 線種      : LinePatternElement の名前一致
    /// で移行先の ID へ解決し直す。解決できなかったものはレポートに残す。
    ///
    /// ⚠️ 書き込みは「書けたか」ではなく「ユーザーが見るビューで効いたか」で判定する。
    /// ビューの V/G が従属ビューやビューテンプレートに支配されていると、
    /// ビュー自身への SetCategoryOverrides は例外になるか黙って無視される。
    /// その場合は主ビュー／テンプレートへ振り替えるが、
    /// 振り替え先へ書けたことは元のビューに反映されたことを意味しない
    /// （テンプレートの V/G に「読み込み」が含まれていなければ伝わらない）。
    /// そのため検証は常に元のビューで行い、伝わらなければ書き込みを破棄する。
    /// </summary>
    public class DwgLayerApplier
    {
        /// <summary>1要素あたり、診断ログに詳細を残すカテゴリ数の上限</summary>
        private const int DetailLogLimit = 5;

        private Dictionary<string, ElementId> _linePatterns;
        private Dictionary<string, ElementId> _fillPatterns;
        private int _detailBudget;

        /// <summary>1要素（ビュー／テンプレート）への書き込み結果。</summary>
        private sealed class WriteStats
        {
            /// <summary>書き込んで実際に効いた件数</summary>
            public int Applied;
            /// <summary>書き込んだが値が変わらなかった件数（テンプレート制御下など）</summary>
            public int Ineffective;
            /// <summary>例外で書き込めなかった件数</summary>
            public int Failed;
            /// <summary>表示/非表示を書き込めなかった件数</summary>
            public int HiddenBlocked;
            /// <summary>最初に起きたエラーの内容（診断用）</summary>
            public string FirstError;

            public int Attempted => Applied + Ineffective + Failed;

            public void Add(WriteStats other)
            {
                Applied += other.Applied;
                Ineffective += other.Ineffective;
                Failed += other.Failed;
                HiddenBlocked += other.HiddenBlocked;
                if (FirstError == null) FirstError = other.FirstError;
            }
        }

        /// <summary>
        /// 移行元 1 DWG 分のレイヤ設定を、移行先の複数ビュー × 複数 DWG へ適用する。
        /// トランザクションは内部で開始・コミットする。
        /// </summary>
        /// <param name="targetDoc">移行先ドキュメント</param>
        /// <param name="sourceLayers">移行元のレイヤ名 -&gt; 設定（"" は DWG 本体）</param>
        /// <param name="targetViews">
        /// 適用先のビュー。チェックされたものをそのまま渡してよい（重複排除は不要）。
        /// 直接書けないビューは、このメソッドがビューテンプレートへ振り替える。
        /// </param>
        /// <param name="targetDwgs">適用先の DWG（複数可）</param>
        public TransferResult Apply(
            Document targetDoc,
            IDictionary<string, LayerGraphicSetting> sourceLayers,
            IList<ViewEntry> targetViews,
            IList<DwgDefinition> targetDwgs,
            IDictionary<string, DwgObjectStyle> sourceStyles)
        {
            var result = new TransferResult();
            if (targetDoc == null || sourceLayers == null || sourceLayers.Count == 0) return result;
            if (targetViews == null || targetViews.Count == 0) return result;
            if (targetDwgs == null || targetDwgs.Count == 0) return result;

            _linePatterns = BuildLinePatternMap(targetDoc);
            _fillPatterns = BuildFillPatternMap(targetDoc);
            result.SourceSettingCount = DwgLayerScanner.CountConfigured(sourceLayers.Values);

            DiagLog.Write($"[DwgVg] ===== 適用開始 ビュー={targetViews.Count} DWG={targetDwgs.Count} " +
                          $"レイヤ設定={sourceLayers.Count} (既定以外={result.SourceSettingCount}) " +
                          $"移行先線種={_linePatterns.Count} 移行先塗潰し={_fillPatterns.Count}");
            foreach (var d in targetDwgs)
                DiagLog.Write($"[DwgVg]   移行先DWG '{d?.Name}' レイヤ数={d?.Layers.Count} catId={d?.CategoryId}");
            DiagLog.Write("[DwgVg]   " + DescribeSource(sourceLayers));

            var missingLayers = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            var missingPatterns = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            // 同じ振り替え先を共有するビューが複数あっても、書き込みは1回で済ませる
            var writtenTargets = new HashSet<ElementId>();

            using (var t = new Transaction(targetDoc, Loc.S("DwgVg.Txn.Apply")))
            {
                t.Start();

                // V/G はオブジェクトスタイルへの差分でしかないため、先に基準値を揃える。
                // これを移さないと、上書きしていないレイヤの見た目がモデル間で揃わない
                if (sourceStyles != null && sourceStyles.Count > 0)
                {
                    foreach (var dwg in targetDwgs)
                        result.ObjectStyleCount += ApplyObjectStyles(targetDoc, sourceStyles, dwg, missingPatterns);
                }

                foreach (var entry in targetViews)
                {
                    if (entry == null) continue;

                    if (!(targetDoc.GetElement(entry.Id) is View view))
                    {
                        result.Failures.Add(string.Format(Loc.S("DwgVg.Fail.ViewNotFound"), entry.Name));
                        continue;
                    }

                    var w = WriteAll(view, entry.Name, sourceLayers, targetDwgs, missingLayers, missingPatterns);
                    DiagLog.Write($"[DwgVg]   '{entry.Name}' 直接: 反映={w.Applied} 未反映={w.Ineffective} " +
                                  $"失敗={w.Failed} {w.FirstError}");

                    result.HiddenBlocked += w.HiddenBlocked;

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

                    // ビュー側では効かなかった。主ビュー／ビューテンプレートへ振り替える
                    ApplyViaFallback(targetDoc, entry, view, sourceLayers, targetDwgs,
                                     missingLayers, missingPatterns, writtenTargets, w, result);
                }

                t.Commit();
            }

            result.MissingLayers.AddRange(missingLayers.OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase));
            result.MissingLinePatterns.AddRange(missingPatterns.OrderBy(s => s, StringComparer.CurrentCultureIgnoreCase));

            DiagLog.Write($"[DwgVg] ===== 適用終了 反映ビュー={result.ViewCount} 反映レイヤ={result.LayerCount} " +
                          $"テンプレート振替={result.TemplateFallbackViews}ビュー/{result.TemplateFallbackCount}件 " +
                          $"未一致レイヤ={result.MissingLayers.Count} 未解決線種={result.MissingLinePatterns.Count} " +
                          $"非表示不可={result.HiddenBlocked} オブジェクトスタイル={result.ObjectStyleCount} " +
                          $"失敗={result.Failures.Count}");
            return result;
        }

        /// <summary>
        /// ビューへ直接書けなかった場合の振り替え先へ順に書き込む。
        ///
        /// ビューの「読み込みカテゴリ」V/G は、次のいずれかに支配されていることがあり、
        /// その場合ビュー自身への SetCategoryOverrides は例外になる。
        ///   1. 従属ビュー  … 主ビューの V/G を継承する
        ///   2. ビューテンプレート … テンプレートが「読み込み」を含めて制御している
        ///
        /// ⚠️ 振り替え先に書けたことは、元のビューに反映されたことを意味しない。
        /// 例えばテンプレートの V/G に「読み込み」が含まれていなければ、
        /// テンプレートへの書き込みは成功してもビューには一切伝わらない。
        /// そのため検証は必ず「元のビュー」で行い、伝わらなかった書き込みは
        /// サブトランザクションごと破棄してモデルを汚さない。
        /// </summary>
        private void ApplyViaFallback(
            Document targetDoc, ViewEntry entry, View view,
            IDictionary<string, LayerGraphicSetting> sourceLayers, IList<DwgDefinition> targetDwgs,
            HashSet<string> missingLayers, HashSet<string> missingPatterns,
            HashSet<ElementId> writtenTargets, WriteStats direct, TransferResult result)
        {
            // ビューテンプレート自身に書いて効かなかった場合は、振り替え先が無い
            if (entry.IsTemplate)
            {
                result.Failures.Add(string.Format(
                    Loc.S("DwgVg.Fail.NotApplied"), entry.Name, direct.FirstError ?? ""));
                return;
            }

            var chain = BuildFallbackChain(targetDoc, view);
            if (chain.Count == 0)
            {
                result.Failures.Add(string.Format(
                    Loc.S("DwgVg.Fail.NotApplied"), entry.Name, direct.FirstError ?? ""));
                return;
            }

            result.TemplateFallbackViews++;
            string lastTarget = "";

            DiagLog.Write($"[DwgVg]   '{entry.Name}' の振替先候補: " +
                          string.Join(" → ", chain.Select(SafeViewName)));

            foreach (var target in chain)
            {
                string targetName;
                try { targetName = target.Name ?? ""; } catch { targetName = ""; }
                lastTarget = targetName;

                // 同じ振り替え先には一度だけ書けばよい。
                // 2件目以降のビューは、書き込み済みの状態で伝播だけ確認する
                if (!writtenTargets.Contains(target.Id))
                {
                    using (var sub = new SubTransaction(targetDoc))
                    {
                        sub.Start();

                        var wt = WriteAll(target, targetName, sourceLayers, targetDwgs,
                                          missingLayers, missingPatterns);
                        try { targetDoc.Regenerate(); } catch { }

                        int ok = CountEffectiveOnView(view, sourceLayers, targetDwgs, out int configured);
                        DiagLog.Write($"[DwgVg]   → '{targetName}' へ振替: 反映={wt.Applied} " +
                                      $"未反映={wt.Ineffective} 失敗={wt.Failed} / " +
                                      $"ビュー '{entry.Name}' への伝播={ok}/{configured} {wt.FirstError}");

                        if (wt.Applied > 0 && (configured == 0 || ok > 0))
                        {
                            sub.Commit();
                            writtenTargets.Add(target.Id);
                            result.TemplateFallbackCount++;
                            result.ViewCount++;
                            result.LayerCount += wt.Applied;
                            return;
                        }

                        // 伝わらなかった書き込みは残さない
                        sub.RollBack();
                    }
                    continue;
                }

                int ok2 = CountEffectiveOnView(view, sourceLayers, targetDwgs, out int configured2);
                DiagLog.Write($"[DwgVg]   → '{targetName}' は書込済み。" +
                              $"ビュー '{entry.Name}' への伝播={ok2}/{configured2}");
                if (configured2 == 0 || ok2 > 0)
                {
                    result.ViewCount++;
                    return;
                }
            }

            result.Failures.Add(string.Format(
                Loc.S("DwgVg.Fail.NoEffect"), entry.Name, lastTarget));
        }

        private static string SafeViewName(View v)
        {
            try { return v?.Name ?? ""; } catch { return "?"; }
        }

        /// <summary>
        /// ビューへ直接書けないときに試す書き込み先を、優先順に並べる。
        /// 従属ビューなら主ビュー、次に（主ビューまたは自身の）ビューテンプレート。
        /// </summary>
        private static List<View> BuildFallbackChain(Document doc, View view)
        {
            var chain = new List<View>();

            View basis = view;
            try
            {
                ElementId primaryId = view.GetPrimaryViewId();
                if (primaryId != null && primaryId != ElementId.InvalidElementId
                    && doc.GetElement(primaryId) is View primary)
                {
                    chain.Add(primary);
                    basis = primary;
                }
            }
            catch { }

            var tpl = DwgLayerScanner.GetAssignedTemplate(doc, basis);
            if (tpl != null) chain.Add(tpl);

            return chain;
        }

        /// <summary>
        /// 指定ビューで、移行元の設定が実際に効いている件数を数える。
        /// 振り替え先へ書いた内容がそのビューへ伝わったかの判定に使う。
        /// </summary>
        /// <param name="configuredCount">移行元で既定から変更されていた設定の件数（判定の母数）</param>
        private int CountEffectiveOnView(
            View view, IDictionary<string, LayerGraphicSetting> sourceLayers,
            IList<DwgDefinition> targetDwgs, out int configuredCount)
        {
            int ok = 0;
            configuredCount = 0;

            foreach (var dwg in targetDwgs)
            {
                if (dwg == null) continue;

                foreach (var kv in sourceLayers)
                {
                    var s = kv.Value;
                    // 既定のままの設定は「伝わったか」の判定に使えない（元から一致するため）
                    if (s == null || !s.HasAnySetting) continue;

                    ElementId catId;
                    if (kv.Key.Length == 0) catId = dwg.CategoryId;
                    else if (!dwg.Layers.TryGetValue(kv.Key, out catId)) continue;

                    configuredCount++;
                    if (IsEffective(view, catId, s)) ok++;
                }
            }

            return ok;
        }

        /// <summary>
        /// DWG のオブジェクトスタイル（モデル全体の基準値）を移行先へ書き込む。
        /// ⚠️ ビュー単位ではなくモデル全体に効く。
        /// </summary>
        /// <returns>書き換えたレイヤ数</returns>
        private int ApplyObjectStyles(
            Document doc, IDictionary<string, DwgObjectStyle> sourceStyles,
            DwgDefinition targetDwg, HashSet<string> missingPatterns)
        {
            var parent = DwgLayerScanner.FindCategory(doc, targetDwg.CategoryId);
            if (parent == null)
            {
                DiagLog.Write($"[DwgVg]   オブジェクトスタイル: '{targetDwg.Name}' のカテゴリを取得できません");
                return 0;
            }

            int applied = 0, failed = 0;
            string firstError = null;

            foreach (var kv in sourceStyles)
            {
                var st = kv.Value;
                if (st == null) continue;

                var cat = DwgLayerScanner.FindSubCategory(parent, kv.Key);
                if (cat == null) continue;

                try
                {
                    if (st.LineColor != null && st.LineColor.IsValid)
                        cat.LineColor = st.LineColor;

                    if (st.ProjectionLineWeight > 0)
                        cat.SetLineWeight(st.ProjectionLineWeight, GraphicsStyleType.Projection);
                    if (st.CutLineWeight > 0)
                        cat.SetLineWeight(st.CutLineWeight, GraphicsStyleType.Cut);

                    var proj = ResolveLinePattern(st.ProjectionLinePattern, missingPatterns);
                    if (proj != null) cat.SetLinePatternId(proj, GraphicsStyleType.Projection);

                    var cut = ResolveLinePattern(st.CutLinePattern, missingPatterns);
                    if (cut != null) cat.SetLinePatternId(cut, GraphicsStyleType.Cut);

                    applied++;
                }
                catch (Exception ex)
                {
                    failed++;
                    if (firstError == null) firstError = ex.Message;
                }
            }

            DiagLog.Write($"[DwgVg]   オブジェクトスタイル '{targetDwg.Name}': 書換={applied} 失敗={failed} {firstError}");
            return applied;
        }

        /// <summary>1要素へ、選択された全 DWG 分の設定を書き込む。</summary>
        private WriteStats WriteAll(
            View view, string viewName,
            IDictionary<string, LayerGraphicSetting> sourceLayers, IList<DwgDefinition> targetDwgs,
            HashSet<string> missingLayers, HashSet<string> missingPatterns)
        {
            var total = new WriteStats();
            _detailBudget = DetailLogLimit;

            foreach (var dwg in targetDwgs)
            {
                if (dwg == null) continue;

                var st = new WriteStats();
                int notFound = 0;

                foreach (var kv in sourceLayers)
                {
                    string layerName = kv.Key;
                    LayerGraphicSetting setting = kv.Value;
                    if (setting == null) continue;

                    // レイヤ名 "" は DWG 本体（親カテゴリ）
                    ElementId targetCatId;
                    if (layerName.Length == 0)
                    {
                        targetCatId = dwg.CategoryId;
                    }
                    else if (!dwg.Layers.TryGetValue(layerName, out targetCatId))
                    {
                        missingLayers.Add(layerName);
                        notFound++;
                        continue;
                    }

                    WriteOne(view, targetCatId, layerName, setting, missingPatterns, st);
                }

                DiagLog.Write($"[DwgVg]     '{viewName}' × '{dwg.Name}': 反映={st.Applied} " +
                              $"未反映={st.Ineffective} 失敗={st.Failed} レイヤ未一致={notFound} {st.FirstError}");
                total.Add(st);
            }

            return total;
        }

        /// <summary>1カテゴリ分の設定を書き込み、実際に効いたかを読み戻して確認する。</summary>
        private void WriteOne(
            View view, ElementId categoryId, string layerName, LayerGraphicSetting s,
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
                LogDetail(view, categoryId, layerName, s, "SetCategoryOverrides 例外: " + ex.Message);
                return;
            }

            // 表示/非表示は別 API。
            // CanCategoryBeHidden で事前に諦めると「非表示にできていない」ことに気付けないため、
            // 実際に書いて例外を拾い、書けなかった件数を残す
            try
            {
                if (view.GetCategoryHidden(categoryId) != s.Hidden)
                    view.SetCategoryHidden(categoryId, s.Hidden);
            }
            catch (Exception ex)
            {
                if (s.Hidden) st.HiddenBlocked++;
                if (st.FirstError == null) st.FirstError = ex.Message;
            }

            bool ok = IsEffective(view, categoryId, s);
            if (ok) st.Applied++;
            else st.Ineffective++;

            // 効かなかったものを優先して詳細を残す（原因の切り分け用）
            if (!ok || s.HasAnySetting) LogDetail(view, categoryId, layerName, s, ok ? "反映" : "未反映");
        }

        /// <summary>書き込んだ値と読み戻した値を診断ログに残す（1要素あたり数件まで）。</summary>
        private void LogDetail(View view, ElementId categoryId, string layerName, LayerGraphicSetting s, string verdict)
        {
            if (_detailBudget <= 0) return;
            _detailBudget--;

            string actual;
            try
            {
                var cur = view.GetCategoryOverrides(categoryId);
                string hidden;
                try { hidden = view.GetCategoryHidden(categoryId).ToString(); }
                catch (Exception ex) { hidden = "取得不可(" + ex.Message + ")"; }

                actual = cur == null
                    ? "上書き取得不可"
                    : $"線色={ColorText(cur.ProjectionLineColor)} 線幅={cur.ProjectionLineWeight} " +
                      $"線種={PatternText(view.Document, cur.ProjectionLinePatternId)} " +
                      $"HT={cur.Halftone} 非表示={hidden}";
            }
            catch (Exception ex) { actual = "読戻し例外: " + ex.Message; }

            DiagLog.Write($"[DwgVg]       [{verdict}] cat={categoryId} layer='{layerName}' " +
                          $"書込(線色={ColorText(s.ProjectionLineColor)} 線幅={s.ProjectionLineWeight} " +
                          $"線種={s.ProjectionLinePattern ?? "-"} HT={s.Halftone} 非表示={s.Hidden}) " +
                          $"実際({actual})");
        }

        private static string ColorText(Color c)
        {
            try { return (c != null && c.IsValid) ? $"{c.Red},{c.Green},{c.Blue}" : "-"; }
            catch { return "?"; }
        }

        private static string PatternText(Document doc, ElementId id)
        {
            try
            {
                if (id == null || id == ElementId.InvalidElementId) return "-";
                if (id == LinePatternElement.GetSolidPatternId()) return LayerGraphicSetting.SolidPatternMarker;
                return (doc?.GetElement(id) as LinePatternElement)?.Name ?? id.ToString();
            }
            catch { return "?"; }
        }

        /// <summary>
        /// 移行元がどんな種類の設定を持っているかを1行にまとめる。
        /// 「移行しても見た目が変わらない」ときに、何が来ているはずなのかを特定するために使う。
        /// </summary>
        private static string DescribeSource(IDictionary<string, LayerGraphicSetting> sourceLayers)
        {
            int hidden = 0, color = 0, weight = 0, pattern = 0, halftone = 0, cut = 0;

            foreach (var s in sourceLayers.Values)
            {
                if (s == null) continue;
                if (s.Hidden) hidden++;
                if (s.ProjectionLineColor != null && s.ProjectionLineColor.IsValid) color++;
                if (s.ProjectionLineWeight > 0) weight++;
                if (s.ProjectionLinePattern != null) pattern++;
                if (s.Halftone) halftone++;
                if ((s.CutLineColor != null && s.CutLineColor.IsValid)
                    || s.CutLineWeight > 0 || s.CutLinePattern != null) cut++;
            }

            return $"移行元内訳: 非表示={hidden} 線色={color} 線幅={weight} 線種={pattern} " +
                   $"ハーフトーン={halftone} 切断線={cut}";
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

            // サーフェス／切断のパターン。以前は「DWG には効かない」と決めつけて除外していたが、
            // 設定が入っている場合に移行しても見た目が変わらない原因になるため必ず移す
            var sfg = ResolveFillPattern(s.SurfaceFgPattern, missingPatterns);
            if (sfg != null) ogs.SetSurfaceForegroundPatternId(sfg);
            if (s.SurfaceFgColor != null && s.SurfaceFgColor.IsValid)
                ogs.SetSurfaceForegroundPatternColor(s.SurfaceFgColor);
            if (!s.SurfaceFgVisible) ogs.SetSurfaceForegroundPatternVisible(false);

            var sbg = ResolveFillPattern(s.SurfaceBgPattern, missingPatterns);
            if (sbg != null) ogs.SetSurfaceBackgroundPatternId(sbg);
            if (s.SurfaceBgColor != null && s.SurfaceBgColor.IsValid)
                ogs.SetSurfaceBackgroundPatternColor(s.SurfaceBgColor);
            if (!s.SurfaceBgVisible) ogs.SetSurfaceBackgroundPatternVisible(false);

            var cfg = ResolveFillPattern(s.CutFgPattern, missingPatterns);
            if (cfg != null) ogs.SetCutForegroundPatternId(cfg);
            if (s.CutFgColor != null && s.CutFgColor.IsValid)
                ogs.SetCutForegroundPatternColor(s.CutFgColor);
            if (!s.CutFgVisible) ogs.SetCutForegroundPatternVisible(false);

            var cbg = ResolveFillPattern(s.CutBgPattern, missingPatterns);
            if (cbg != null) ogs.SetCutBackgroundPatternId(cbg);
            if (s.CutBgColor != null && s.CutBgColor.IsValid)
                ogs.SetCutBackgroundPatternColor(s.CutBgColor);
            if (!s.CutBgVisible) ogs.SetCutBackgroundPatternVisible(false);

            if (s.Transparency > 0) ogs.SetSurfaceTransparency(s.Transparency);
            if (s.DetailLevel != ViewDetailLevel.Undefined) ogs.SetDetailLevel(s.DetailLevel);

            ogs.SetHalftone(s.Halftone);
            return ogs;
        }

        /// <summary>塗潰しパターン名を移行先の ElementId へ解決する。</summary>
        private ElementId ResolveFillPattern(string name, HashSet<string> missingPatterns)
        {
            if (name == null) return null;
            if (_fillPatterns.TryGetValue(name, out var id)) return id;

            missingPatterns.Add(name);
            return null;
        }

        private static Dictionary<string, ElementId> BuildFillPatternMap(Document doc)
        {
            var map = new Dictionary<string, ElementId>(StringComparer.CurrentCultureIgnoreCase);
            try
            {
                foreach (var e in new FilteredElementCollector(doc)
                             .OfClass(typeof(FillPatternElement))
                             .ToElements())
                {
                    if (e is FillPatternElement fpe && !string.IsNullOrEmpty(fpe.Name))
                        map[fpe.Name] = fpe.Id;
                }
            }
            catch { }
            return map;
        }

        /// <summary>
        /// 書き込んだ設定がそのビューで実際に効いているかを判定する。
        ///
        /// ビューテンプレートに制御されているビューでは GetCategoryOverrides が
        /// テンプレート側の値を返すため、書いた値と食い違えば「効いていない」と分かる。
        /// 線種は移行先に同名が無いと書き込まない（別途レポートする）ため比較対象から外す。
        /// </summary>
        private bool IsEffective(View view, ElementId categoryId, LayerGraphicSetting s)
        {
            try
            {
                // 表示/非表示は CanCategoryBeHidden で判定を省かない。
                // 省くと「非表示にできていない」まま成功と報告してしまう
                try
                {
                    if (view.GetCategoryHidden(categoryId) != s.Hidden) return false;
                }
                catch
                {
                    if (s.Hidden) return false;   // 非表示にしたいのに状態が読めない＝再現できていない
                }

                var cur = view.GetCategoryOverrides(categoryId);
                if (cur == null) return false;

                if (cur.Halftone != s.Halftone) return false;

                if (!SameColor(cur.ProjectionLineColor, s.ProjectionLineColor)) return false;
                if (!SameColor(cur.CutLineColor, s.CutLineColor)) return false;

                if (cur.ProjectionLineWeight != (s.ProjectionLineWeight > 0 ? s.ProjectionLineWeight : -1))
                    return false;
                if (cur.CutLineWeight != (s.CutLineWeight > 0 ? s.CutLineWeight : -1))
                    return false;

                if (cur.Transparency != s.Transparency) return false;
                if (cur.DetailLevel != s.DetailLevel) return false;

                // 線種・塗潰しは移行先に同名が無ければ書き込んでいないので、解決できたものだけ照合する
                if (!SamePattern(cur.ProjectionLinePatternId, s.ProjectionLinePattern)) return false;
                if (!SamePattern(cur.CutLinePatternId, s.CutLinePattern)) return false;

                if (!SameFill(cur.SurfaceForegroundPatternId, s.SurfaceFgPattern)) return false;
                if (!SameFill(cur.SurfaceBackgroundPatternId, s.SurfaceBgPattern)) return false;
                if (!SameFill(cur.CutForegroundPatternId, s.CutFgPattern)) return false;
                if (!SameFill(cur.CutBackgroundPatternId, s.CutBgPattern)) return false;

                if (!SameColor(cur.SurfaceForegroundPatternColor, s.SurfaceFgColor)) return false;
                if (!SameColor(cur.SurfaceBackgroundPatternColor, s.SurfaceBgColor)) return false;
                if (!SameColor(cur.CutForegroundPatternColor, s.CutFgColor)) return false;
                if (!SameColor(cur.CutBackgroundPatternColor, s.CutBgColor)) return false;

                if (cur.IsSurfaceForegroundPatternVisible != s.SurfaceFgVisible) return false;
                if (cur.IsSurfaceBackgroundPatternVisible != s.SurfaceBgVisible) return false;
                if (cur.IsCutForegroundPatternVisible != s.CutFgVisible) return false;
                if (cur.IsCutBackgroundPatternVisible != s.CutBgVisible) return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 線種の一致判定。移行先で解決できなかった線種は書き込んでいないため、一致とみなす
        /// （未解決分は MissingLinePatterns で別途報告する）。
        /// </summary>
        private bool SamePattern(ElementId current, string wantedName)
        {
            if (wantedName == null) return true;   // 上書きなし（元の値を消したかどうかは他項目で判定）

            ElementId wanted;
            if (wantedName == LayerGraphicSetting.SolidPatternMarker)
            {
                try { wanted = LinePatternElement.GetSolidPatternId(); }
                catch { return true; }
            }
            else if (!_linePatterns.TryGetValue(wantedName, out wanted))
            {
                return true;   // 移行先に存在しない線種
            }

            return current != null && current == wanted;
        }

        /// <summary>塗潰しパターンの一致判定。移行先に無いものは書き込んでいないため一致とみなす。</summary>
        private bool SameFill(ElementId current, string wantedName)
        {
            if (wantedName == null) return true;
            if (!_fillPatterns.TryGetValue(wantedName, out var wanted)) return true;
            return current != null && current == wanted;
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
