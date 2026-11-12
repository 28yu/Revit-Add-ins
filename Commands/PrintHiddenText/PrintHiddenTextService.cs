using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Tools28.Localization;

namespace Tools28.Commands.PrintHiddenText
{
    /// <summary>
    /// 「画面には出るが、印刷・PDF には出ないテキスト」を実現する常駐処理。
    ///
    /// タイプ名に「非表示」を含むテキストノート（TextNote）と寸法（Dimension。スポット寸法を含む）を、
    ///   印刷:  DocumentPrinting で隠す → DocumentPrinted で戻す
    ///   PDF:   FileExporting   で隠す → FileExported   で戻す（形式が PDF のときだけ）
    /// という流れで、出力の間だけビューから隠す（View.HideElements / UnhideElements）。
    /// 要素の削除・作り直しはしない。ボタンは無く、OnStartup で登録すれば自動で働く。
    ///
    /// 【戻せなかったときの安全策】
    /// 出力が失敗・取り消しされた直後はモデルを変更できないことがある（API の仕様）。
    /// そのときは隠した要素の ID を覚えたまま、Revit が空いたとき（Idling）に戻し直す。
    /// 何度試しても戻せない場合は、どのビューで残っているかを画面で知らせる。
    ///
    /// 【カーテンパネルのネスト注釈（2026-11-12 追加）】
    /// カーテンパネルのタイプパラメータ「幅数値表示」（Yes/No）を出力の間だけ 0 にし、
    /// 出力後に元の値（1）へ戻す。ネストした一般注釈の表示をこのパラメータに関連付けておけば、
    /// 画面には数値が出て、印刷・PDF には出なくなる。テキストの処理とは独立していて、
    /// どちらか一方だけが対象のモデルでも動く。
    ///
    /// 【TransactionGroup について】
    /// Revit は「イベントの処理を終える時点で、開いているトランザクション（グループを含む）は
    /// すべて閉じていなければならない」。そのため出力前に開いたグループを出力後まで持ち越して
    /// RollBack することはできない。代わりに、出力前・出力後のそれぞれで
    /// TransactionGroup にまとめて Assimilate し（元に戻す履歴が 1 件ずつになる）、
    /// 出力後は「変更前の値」の記録から確実に戻す。途中で失敗したらグループごと RollBack する。
    /// </summary>
    internal sealed class PrintHiddenTextService
    {
        public static PrintHiddenTextService Instance { get; } = new PrintHiddenTextService();

        /// <summary>
        /// 対象のテキストタイプを見分ける文字列。
        /// ⚠ 利用者がタイプ名に付ける「検索キー」なので多言語化しない（CLAUDE.md の分類 C）。
        /// </summary>
        internal const string Keyword = "非表示";

        /// <summary>
        /// カーテンパネルのタイプパラメータ名（Yes/No）。0 にするとネスト注釈が消える前提。
        /// ⚠ ファミリ側のパラメータ名を探すキーなので多言語化しない（CLAUDE.md の分類 C）。
        /// </summary>
        internal const string PanelParamName = "幅数値表示";

        /// <summary>Idling での戻し直しをあきらめて利用者に知らせるまでの回数</summary>
        private const int MaxRestoreAttempts = 30;

        private UIControlledApplication _app;
        private bool _idlingHooked;

        /// <summary>今回の印刷で DocumentPrinting が処理を終えたか（終えていればビューごとの予備処理はしない）</summary>
        private bool _printHandled;

        /// <summary>モデルごとの「隠したもの」の記録。</summary>
        private readonly List<HiddenRecord> _records = new List<HiddenRecord>();

        private sealed class HiddenRecord
        {
            public Document Doc;
            /// <summary>ビュー → そのビューで隠した要素</summary>
            public readonly Dictionary<ElementId, HashSet<ElementId>> ByView = new Dictionary<ElementId, HashSet<ElementId>>();
            /// <summary>カーテンパネルのタイプ → 「幅数値表示」の変更前の値</summary>
            public readonly Dictionary<ElementId, int> PanelTypeValues = new Dictionary<ElementId, int>();
            public int Attempts;
            public bool IsEmpty => ByView.Count == 0 && PanelTypeValues.Count == 0;
        }

        private PrintHiddenTextService() { }

        // ================================================================
        // 登録・解除
        // ================================================================

        public void Initialize(UIControlledApplication app)
        {
            try
            {
                _app = app;
                var ca = app.ControlledApplication;
                ca.DocumentPrinting += OnDocumentPrinting;
                ca.ViewPrinting += OnViewPrinting;
                ca.DocumentPrinted += OnDocumentPrinted;
                ca.FileExporting += OnFileExporting;
                ca.FileExported += OnFileExported;
                ca.DocumentClosing += OnDocumentClosing;
                DiagLog.Write("PrintHiddenText 初期化完了");
            }
            catch (Exception ex)
            {
                DiagLog.Write($"PrintHiddenText 初期化失敗: {ex.Message}");
            }
        }

        public void Shutdown()
        {
            if (_app == null) return;
            try
            {
                var ca = _app.ControlledApplication;
                ca.DocumentPrinting -= OnDocumentPrinting;
                ca.ViewPrinting -= OnViewPrinting;
                ca.DocumentPrinted -= OnDocumentPrinted;
                ca.FileExporting -= OnFileExporting;
                ca.FileExported -= OnFileExported;
                ca.DocumentClosing -= OnDocumentClosing;
                UnhookIdling();
            }
            catch { }
            _app = null;
        }

        // ================================================================
        // イベント
        // ================================================================

        private void OnDocumentPrinting(object sender, DocumentPrintingEventArgs e)
        {
            try
            {
                _printHandled = false;
                var viewIds = e.GetViewElementIds();
                DiagLog.Write($"PrintHiddenText 印刷前: ビュー {viewIds?.Count ?? 0} 件");
                _printHandled = PrepareForOutput(e.Document, viewIds, "印刷前");
            }
            catch (Exception ex)
            {
                // ここで例外を外に出すと印刷そのものを邪魔するので、記録だけして通常どおり出力させる
                DiagLog.Write($"PrintHiddenText 印刷前の処理で例外: {ex}");
            }
        }

        /// <summary>
        /// 予備: ビュー1枚ごとの印刷直前。DocumentPrinting で隠せなかった場合だけ、
        /// このビュー（シートなら配置ビューも）についてもう一度隠す。戻すのは DocumentPrinted でまとめて行う。
        /// </summary>
        private void OnViewPrinting(object sender, ViewPrintingEventArgs e)
        {
            try
            {
                var view = e.View;
                var doc = view?.Document;
                if (doc == null || _printHandled) return;
                DiagLog.Write($"PrintHiddenText ビュー印刷前（予備）: 「{SafeName(view)}」 {e.Index + 1}/{e.TotalViews}");
                PrepareForOutput(doc, new List<ElementId> { view.Id }, "ビュー印刷前（予備）");
            }
            catch (Exception ex)
            {
                DiagLog.Write($"PrintHiddenText ビュー印刷前の処理で例外: {ex}");
            }
        }

        private void OnDocumentPrinted(object sender, DocumentPrintedEventArgs e)
        {
            _printHandled = false;
            AfterOutput(e.Document, "印刷後");
        }

        private void OnFileExporting(object sender, FileExportingEventArgs e)
        {
            try
            {
                string format = SafeFormat(e);
                DiagLog.Write($"PrintHiddenText 書き出し前: 形式={format}, 保存先={e.Path}");
                if (!IsPdf(format, e.Path)) return;   // DWG などには一切手を出さない

                // FileExporting では出力するビューの一覧が取れないため、
                // 対象テキストが見えるビューをモデル全体から探して隠す
                PrepareForOutput(e.Document, null, "PDF書き出し前");
            }
            catch (Exception ex)
            {
                DiagLog.Write($"PrintHiddenText 書き出し前の処理で例外: {ex}");
            }
        }

        private void OnFileExported(object sender, FileExportedEventArgs e)
        {
            // 記録があるのは PDF のときに隠した場合だけなので、形式を問わず「記録があれば戻す」
            AfterOutput(e.Document, "書き出し後");
        }

        private void OnDocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            try
            {
                var rec = FindRecord(e.Document);
                if (rec == null) return;
                // 閉じる途中のモデルは変更しない（変更すると保存の確認が再び出るなどの副作用があるため）。
                // 出力中は保存できないので、隠した状態がファイルに残ることはない。記録だけ破棄する
                DiagLog.Write("PrintHiddenText 戻す前にモデルが閉じられたため、記録を破棄しました");
                _records.Remove(rec);
                if (_records.Count == 0) UnhookIdling();
            }
            catch (Exception ex)
            {
                DiagLog.Write($"PrintHiddenText モデルを閉じる前の処理で例外: {ex.Message}");
            }
        }

        private void AfterOutput(Document doc, string phase)
        {
            try
            {
                var rec = FindRecord(doc);
                if (rec == null) return;
                if (TryRestore(rec))
                {
                    DiagLog.Write($"PrintHiddenText {phase}: 元に戻しました");
                    return;
                }
                // 出力が失敗・取り消しされた直後などは変更できないので、空いたときにやり直す
                DiagLog.Write($"PrintHiddenText {phase}: すぐには戻せないため、Revit が空いたときに戻します");
                HookIdling();
            }
            catch (Exception ex)
            {
                DiagLog.Write($"PrintHiddenText {phase}の処理で例外: {ex}");
                HookIdling();
            }
        }

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            try
            {
                foreach (var rec in _records.ToList())
                {
                    if (rec.Doc == null || !rec.Doc.IsValidObject)
                    {
                        _records.Remove(rec);
                        continue;
                    }
                    // 編集中などで変更できないあいだは数えない（空くまで待つ）
                    if (!CanModify(rec.Doc)) continue;
                    if (TryRestore(rec)) continue;

                    rec.Attempts++;
                    if (rec.Attempts >= MaxRestoreAttempts)
                    {
                        // 戻せないまま黙っていると、テキストが消えたと誤解されるので知らせる
                        string panelTypes = string.Join("\n", rec.PanelTypeValues.Keys
                            .Select(id => SafeTypeName(rec.Doc.GetElement(id)))
                            .Where(n => !string.IsNullOrEmpty(n))
                            .Distinct()
                            .Take(20)
                            .Select(n => "・" + n));
                        string views = string.Join("\n", rec.ByView.Keys
                            .Select(id => (rec.Doc.GetElement(id) as View)?.Name)
                            .Where(n => !string.IsNullOrEmpty(n))
                            .Distinct()
                            .Take(20)
                            .Select(n => "・" + n));
                        DiagLog.Write($"PrintHiddenText 戻せなかったビュー:\n{views}\n戻せなかったパネルタイプ:\n{panelTypes}");
                        _records.Remove(rec);
                        if (views.Length > 0)
                            TaskDialog.Show(Loc.S("PrintHiddenText.Title"),
                                string.Format(Loc.S("PrintHiddenText.RestoreFailed"), views));
                        if (panelTypes.Length > 0)
                            TaskDialog.Show(Loc.S("PrintHiddenText.Title"),
                                string.Format(Loc.S("PrintHiddenText.PanelRestoreFailed"), PanelParamName, panelTypes));
                    }
                }
            }
            catch (Exception ex)
            {
                DiagLog.Write($"PrintHiddenText Idling での戻し直しで例外: {ex.Message}");
            }
            finally
            {
                if (_records.Count == 0) UnhookIdling();
            }
        }

        private void HookIdling()
        {
            if (_idlingHooked || _app == null) return;
            _app.Idling += OnIdling;
            _idlingHooked = true;
        }

        private void UnhookIdling()
        {
            if (!_idlingHooked || _app == null) return;
            try { _app.Idling -= OnIdling; } catch { }
            _idlingHooked = false;
        }

        // ================================================================
        // 隠す
        // ================================================================

        /// <param name="viewIds">出力するビュー。null ならモデル全体から対象が見えるビューを探す（PDF 書き出し用）</param>
        /// <param name="phase">ログ用の呼び出し元</param>
        /// <returns>対象が無かった、または変更できたら true。対象があるのに何も変更できなければ false</returns>
        private bool PrepareForOutput(Document doc, ICollection<ElementId> viewIds, string phase)
        {
            DiagLog.Write($"PrintHiddenText {phase}: モデルの状態 {DescribeState(doc)}");
            if (!CanModify(doc))
            {
                DiagLog.Write($"PrintHiddenText {phase}: このモデルは変更できないため何もしません");
                return false;
            }

            // 前回戻せなかった分があれば、先に戻しておく（記録を重ねないため）
            var old = FindRecord(doc);
            if (old != null) TryRestore(old);

            // 何を変えるかを先に決める（ここではモデルを変更しない）
            var annoPlan = BuildAnnotationPlan(doc, viewIds);
            var panelPlan = BuildPanelPlan(doc);
            DiagLog.Write($"PrintHiddenText {phase}: 対象 注釈 {annoPlan.Sum(x => x.Value.Count)} 個（{annoPlan.Count} ビュー）／ カーテンパネル {panelPlan.Count} タイプ");
            if (annoPlan.Count == 0 && panelPlan.Count == 0) return true;   // 対象が無ければ何もしない（通常どおり出力）

            var hiddenNow = new Dictionary<ElementId, HashSet<ElementId>>();
            var panelNow = new Dictionary<ElementId, int>();

            // 印刷の直前は、Revit 自身のトランザクションが開いたまま（IsModifiable = true）のことがある。
            // そのときは新しい Transaction / TransactionGroup を始められないので、SubTransaction で変更する。
            // それ以外は注釈とパネルの変更を 1 つのグループにまとめる（元に戻す履歴は 1 件）。
            // どこかで例外が出たらグループごと RollBack し、何も変えずに通常どおり出力させる
            bool useGroup = !doc.IsModifiable;
            TransactionGroup tg = useGroup ? new TransactionGroup(doc, Loc.S("PrintHiddenText.TxGroupHide")) : null;
            try
            {
                tg?.Start();
                if (annoPlan.Count > 0) HideAnnotations(doc, annoPlan, hiddenNow);
                if (panelPlan.Count > 0) SwitchPanelsOff(doc, panelPlan, panelNow);

                if (tg != null)
                {
                    if (hiddenNow.Count == 0 && panelNow.Count == 0)
                    {
                        tg.RollBack();
                        DiagLog.Write($"PrintHiddenText {phase}: 何も変更できませんでした");
                        return false;
                    }
                    if (tg.Assimilate() != TransactionStatus.Committed)
                    {
                        DiagLog.Write($"PrintHiddenText {phase}: 変更を確定できませんでした（通常どおり出力されます）");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                if (tg != null && tg.GetStatus() == TransactionStatus.Started) tg.RollBack();
                DiagLog.Write($"PrintHiddenText {phase}: 例外（通常どおり出力されます）: {ex}");
                return false;
            }
            finally
            {
                tg?.Dispose();
            }

            if (hiddenNow.Count == 0 && panelNow.Count == 0)
            {
                DiagLog.Write($"PrintHiddenText {phase}: 何も変更できませんでした");
                return false;
            }

            // 確定できたものだけ記録する（これが後で戻すときの唯一の情報源）
            var rec = FindRecord(doc) ?? new HiddenRecord { Doc = doc };
            foreach (var kv in hiddenNow)
            {
                if (!rec.ByView.TryGetValue(kv.Key, out var set))
                    rec.ByView[kv.Key] = set = new HashSet<ElementId>();
                set.UnionWith(kv.Value);
            }
            foreach (var kv in panelNow)
            {
                // 先に記録した「変更前の値」を優先する（0 にした後の値で上書きしない）
                if (!rec.PanelTypeValues.ContainsKey(kv.Key)) rec.PanelTypeValues[kv.Key] = kv.Value;
            }
            rec.Attempts = 0;
            if (!_records.Contains(rec)) _records.Add(rec);

            DiagLog.Write($"PrintHiddenText {phase}: 注釈 {hiddenNow.Count} ビューで計 {hiddenNow.Sum(x => x.Value.Count)} 個を隠し、" +
                          $"カーテンパネル {panelNow.Count} タイプの「{PanelParamName}」を 0 にしました" +
                          (useGroup ? "" : "（Revit のトランザクション内で変更）"));
            return true;
        }

        /// <summary>
        /// 変更を 1 ステップ実行する。モデルにトランザクションが開いていれば SubTransaction、
        /// 開いていなければ Transaction で囲む。body は「何か変更したか」を返す。確定できたら true。
        /// body 内の例外は外へ出す（呼び出し側でグループごと取り消すため）。
        /// </summary>
        private static bool RunStep(Document doc, string name, Func<bool> body)
        {
            if (doc.IsModifiable)
            {
                using (var st = new SubTransaction(doc))
                {
                    st.Start();
                    if (!body()) { st.RollBack(); return false; }
                    return st.Commit() == TransactionStatus.Committed;
                }
            }
            using (var t = new Transaction(doc, name))
            {
                SetQuietFailureHandling(t);
                t.Start();
                if (!body()) { t.RollBack(); return false; }
                return t.Commit() == TransactionStatus.Committed;
            }
        }

        // ---------------- 注釈（テキスト・寸法） ----------------

        /// <summary>
        /// タイプ名に「非表示」を含むテキストタイプ・寸法タイプ（スポット寸法を含む）。
        /// </summary>
        private static HashSet<ElementId> FindTargetTypeIds(Document doc)
        {
            return new HashSet<ElementId>(
                new FilteredElementCollector(doc)
                    .WhereElementIsElementType()
                    .Where(t => (t is TextNoteType || t is DimensionType) && IsTargetTypeName(SafeElementName(t)))
                    .Select(t => t.Id));
        }

        private List<KeyValuePair<View, List<ElementId>>> BuildAnnotationPlan(Document doc, ICollection<ElementId> viewIds)
        {
            var plan = new List<KeyValuePair<View, List<ElementId>>>();

            var typeIds = FindTargetTypeIds(doc);
            if (typeIds.Count == 0) return plan;

            var candidateViews = viewIds != null
                ? ExpandPrintedViews(doc, viewIds)
                : FindViewsShowingTargets(doc, typeIds);

            foreach (var view in candidateViews)
            {
                var ids = CollectVisibleTargets(doc, view, typeIds);
                if (ids.Count > 0) plan.Add(new KeyValuePair<View, List<ElementId>>(view, ids));
            }
            return plan;
        }

        /// <summary>注釈を隠す。ビューごとに SubTransaction にして、1 つのビューで失敗しても他は続ける。</summary>
        private static void HideAnnotations(Document doc, List<KeyValuePair<View, List<ElementId>>> plan,
            Dictionary<ElementId, HashSet<ElementId>> hiddenNow)
        {
            var done = new Dictionary<ElementId, HashSet<ElementId>>();
            bool ok = RunStep(doc, Loc.S("PrintHiddenText.TxHide"), () =>
            {
                foreach (var kv in plan)
                {
                    using (var st = new SubTransaction(doc))
                    {
                        try
                        {
                            st.Start();
                            kv.Key.HideElements(kv.Value);
                            st.Commit();
                            done[kv.Key.Id] = new HashSet<ElementId>(kv.Value);
                        }
                        catch (Exception ex)
                        {
                            if (st.GetStatus() == TransactionStatus.Started) st.RollBack();
                            DiagLog.Write($"PrintHiddenText ビュー「{SafeName(kv.Key)}」で隠せませんでした: {ex.Message}");
                        }
                    }
                }
                return done.Count > 0;
            });
            if (!ok) return;
            foreach (var kv in done) hiddenNow[kv.Key] = kv.Value;
        }

        // ---------------- カーテンパネル ----------------

        /// <summary>
        /// 「幅数値表示」を 0 にするカーテンパネルのタイプと、その変更前の値。
        /// パラメータが無い・読み取り専用・Yes/No（整数）でない・すでに 0・他の人が編集中のタイプは除く。
        /// </summary>
        private static Dictionary<ElementId, int> BuildPanelPlan(Document doc)
        {
            var plan = new Dictionary<ElementId, int>();
            IEnumerable<Element> types;
            try
            {
                types = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                    .WhereElementIsElementType()
                    .ToElements();
            }
            catch { return plan; }

            foreach (var type in types)
            {
                try
                {
                    var p = type.LookupParameter(PanelParamName);
                    if (p == null || p.IsReadOnly || p.StorageType != StorageType.Integer) continue;
                    if (!p.HasValue) continue;
                    int value = p.AsInteger();
                    if (value == 0) continue;   // すでにオフ（変える必要も、戻す必要もない）
                    if (IsOwnedByOther(doc, type.Id)) continue;
                    plan[type.Id] = value;
                }
                catch { /* 読めないタイプは対象外 */ }
            }
            return plan;
        }

        private static void SwitchPanelsOff(Document doc, Dictionary<ElementId, int> plan, Dictionary<ElementId, int> panelNow)
        {
            var done = new Dictionary<ElementId, int>();
            bool ok = RunStep(doc, Loc.S("PrintHiddenText.TxPanelOff"), () =>
            {
                foreach (var kv in plan)
                {
                    try
                    {
                        var p = doc.GetElement(kv.Key)?.LookupParameter(PanelParamName);
                        if (p == null || p.IsReadOnly) continue;
                        if (p.Set(0)) done[kv.Key] = kv.Value;
                    }
                    catch (Exception ex)
                    {
                        DiagLog.Write($"PrintHiddenText パネルタイプ {kv.Key} の「{PanelParamName}」を変更できませんでした: {ex.Message}");
                    }
                }
                return done.Count > 0;
            });
            if (!ok) return;
            foreach (var kv in done) panelNow[kv.Key] = kv.Value;
        }

        private static bool IsOwnedByOther(Document doc, ElementId id)
        {
            try
            {
                return doc.IsWorkshared
                    && WorksharingUtils.GetCheckoutStatus(doc, id) == CheckoutStatus.OwnedByOtherUser;
            }
            catch { return false; }
        }

        // ---------------- ビューの洗い出し ----------------

        /// <summary>印刷するビュー一覧を、シートに配置されたビュー（凡例を含む）まで広げる。</summary>
        private static List<View> ExpandPrintedViews(Document doc, ICollection<ElementId> viewIds)
        {
            var result = new List<View>();
            var seen = new HashSet<ElementId>();

            void Add(View v)
            {
                if (v == null || v.IsTemplate || !seen.Add(v.Id)) return;
                result.Add(v);
            }

            foreach (var id in viewIds)
            {
                var v = doc.GetElement(id) as View;
                Add(v);
                if (v is ViewSheet sheet)
                {
                    foreach (var placedId in sheet.GetAllPlacedViews())
                        Add(doc.GetElement(placedId) as View);
                }
            }
            return result;
        }

        /// <summary>
        /// 出力ビューが分からない場合（PDF 書き出し）に、対象の注釈が見える可能性のあるビューを集める。
        /// テキスト・寸法はビュー固有の要素なので「所属ビュー」で見える。
        /// 従属ビューは親ビューの注釈を共有するので、親と従属ビューもあわせて調べる。
        /// </summary>
        private static List<View> FindViewsShowingTargets(Document doc, HashSet<ElementId> typeIds)
        {
            var ownerIds = new HashSet<ElementId>(
                new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(e => typeIds.Contains(e.GetTypeId()))
                    .Select(e => e.OwnerViewId)
                    .Where(id => id != null && id != ElementId.InvalidElementId));

            var result = new List<View>();
            var seen = new HashSet<ElementId>();

            void Add(ElementId id)
            {
                if (id == null || id == ElementId.InvalidElementId || !seen.Add(id)) return;
                if (doc.GetElement(id) is View v && !v.IsTemplate) result.Add(v);
            }

            foreach (var ownerId in ownerIds)
            {
                Add(ownerId);
                if (!(doc.GetElement(ownerId) is View owner)) continue;
                try
                {
                    var primaryId = owner.GetPrimaryViewId();
                    var root = (primaryId != null && primaryId != ElementId.InvalidElementId)
                        ? doc.GetElement(primaryId) as View
                        : owner;
                    if (root == null) continue;
                    Add(root.Id);
                    foreach (var depId in root.GetDependentViewIds()) Add(depId);
                }
                catch { /* 従属ビューを持てない種類のビュー */ }
            }
            return result;
        }

        /// <summary>ビューで今見えている対象の注釈のうち、隠せるものだけを返す。</summary>
        private static List<ElementId> CollectVisibleTargets(Document doc, View view, HashSet<ElementId> typeIds)
        {
            var ids = new List<ElementId>();
            try
            {
                // ビューを指定した収集は「そのビューで表示されている要素」だけを返す
                // （すでに非表示の要素は入らない＝元から隠れていた要素を後で表示してしまうことがない）
                foreach (var e in new FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType())
                {
                    if (!typeIds.Contains(e.GetTypeId())) continue;
                    if (e.IsHidden(view)) continue;
                    if (!e.CanBeHidden(view)) continue;
                    ids.Add(e.Id);
                }
            }
            catch
            {
                // 集計表など、要素を取り出せない種類のビューは対象外
            }
            return ids;
        }

        // ================================================================
        // 戻す
        // ================================================================

        /// <returns>すべて戻せた（または戻すものが無くなった）ら true</returns>
        private bool TryRestore(HiddenRecord rec)
        {
            var doc = rec.Doc;
            if (doc == null || !doc.IsValidObject || rec.IsEmpty)
            {
                _records.Remove(rec);
                return true;
            }
            if (!CanModify(doc)) return false;

            var viewsDone = new List<ElementId>();
            var panelsDone = new List<ElementId>();

            bool useGroup = !doc.IsModifiable;
            TransactionGroup tg = useGroup ? new TransactionGroup(doc, Loc.S("PrintHiddenText.TxGroupRestore")) : null;
            try
            {
                tg?.Start();
                if (rec.ByView.Count > 0) RestoreAnnotations(doc, rec, viewsDone);
                if (rec.PanelTypeValues.Count > 0) RestorePanels(doc, rec, panelsDone);

                if (tg != null && tg.Assimilate() != TransactionStatus.Committed)
                {
                    DiagLog.Write("PrintHiddenText 戻す処理を確定できませんでした");
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (tg != null && tg.GetStatus() == TransactionStatus.Started) tg.RollBack();
                DiagLog.Write($"PrintHiddenText 戻す処理で例外: {ex.Message}");
                return false;
            }
            finally
            {
                tg?.Dispose();
            }

            foreach (var id in viewsDone) rec.ByView.Remove(id);
            foreach (var id in panelsDone) rec.PanelTypeValues.Remove(id);
            if (rec.IsEmpty)
            {
                _records.Remove(rec);
                return true;
            }
            return false;
        }

        private static void RestoreAnnotations(Document doc, HiddenRecord rec, List<ElementId> viewsDone)
        {
            var done = new List<ElementId>();
            // 何も変更しない（すべて戻し済み）場合も「済み」にするため、変更の有無とは別に数える
            var nothingToDo = new List<ElementId>();
            bool ok = RunStep(doc, Loc.S("PrintHiddenText.TxRestore"), () =>
            {
                foreach (var kv in rec.ByView)
                {
                    var view = doc.GetElement(kv.Key) as View;
                    if (view == null) { nothingToDo.Add(kv.Key); continue; }   // ビューが消えていれば戻す必要なし

                    // 出力の間に削除された要素や、すでに表示に戻っている要素は除く
                    var ids = kv.Value
                        .Where(id => doc.GetElement(id) is Element el && el.IsHidden(view))
                        .ToList();
                    if (ids.Count == 0) { nothingToDo.Add(kv.Key); continue; }

                    using (var st = new SubTransaction(doc))
                    {
                        try
                        {
                            st.Start();
                            view.UnhideElements(ids);
                            st.Commit();
                            done.Add(kv.Key);
                        }
                        catch (Exception ex)
                        {
                            if (st.GetStatus() == TransactionStatus.Started) st.RollBack();
                            DiagLog.Write($"PrintHiddenText ビュー「{SafeName(view)}」で戻せませんでした: {ex.Message}");
                        }
                    }
                }
                return done.Count > 0;
            });
            viewsDone.AddRange(nothingToDo);
            if (ok) viewsDone.AddRange(done);
        }

        private static void RestorePanels(Document doc, HiddenRecord rec, List<ElementId> panelsDone)
        {
            var done = new List<ElementId>();
            var nothingToDo = new List<ElementId>();
            bool ok = RunStep(doc, Loc.S("PrintHiddenText.TxPanelRestore"), () =>
            {
                foreach (var kv in rec.PanelTypeValues)
                {
                    try
                    {
                        var p = doc.GetElement(kv.Key)?.LookupParameter(PanelParamName);
                        // タイプが消えた・パラメータが無くなった場合は戻すものが無い
                        if (p == null || p.StorageType != StorageType.Integer) { nothingToDo.Add(kv.Key); continue; }
                        // すでに 0 以外（Revit 側で取り消された・誰かが変えた）なら、その値を尊重して触らない
                        if (p.AsInteger() != 0) { nothingToDo.Add(kv.Key); continue; }
                        if (p.IsReadOnly) continue;
                        if (p.Set(kv.Value)) done.Add(kv.Key);
                    }
                    catch (Exception ex)
                    {
                        DiagLog.Write($"PrintHiddenText パネルタイプ {kv.Key} の「{PanelParamName}」を戻せませんでした: {ex.Message}");
                    }
                }
                return done.Count > 0;
            });
            panelsDone.AddRange(nothingToDo);
            if (ok) panelsDone.AddRange(done);
        }

        // ================================================================
        // 小物
        // ================================================================

        private HiddenRecord FindRecord(Document doc)
        {
            if (doc == null) return null;
            foreach (var r in _records)
            {
                try
                {
                    if (r.Doc != null && r.Doc.IsValidObject && r.Doc.Equals(doc)) return r;
                }
                catch { }
            }
            return null;
        }

        private static bool CanModify(Document doc)
        {
            try
            {
                return doc != null && doc.IsValidObject
                    && !doc.IsReadOnly
                    && !doc.IsLinked;
                // ⚠ IsModifiable（トランザクションが開いている）は除外しない。印刷の直前は Revit 自身の
                //   トランザクションが開いていることがあり、v1 ではここで弾いて印刷で何も隠れなかった。
                //   開いているときは RunStep が SubTransaction で変更する
            }
            catch { return false; }
        }

        /// <summary>
        /// タイプ名に「非表示」を含むか。全角/半角などの表記ゆれ（NFKC）と大文字小文字を吸収して比べる。
        /// </summary>
        internal static bool IsTargetTypeName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            return Normalize(name).IndexOf(NormalizedKeyword, StringComparison.Ordinal) >= 0;
        }

        private static readonly string NormalizedKeyword = Normalize(Keyword);

        private static string Normalize(string s)
        {
            try { return s.Normalize(NormalizationForm.FormKC).ToUpperInvariant(); }
            catch { return s.ToUpperInvariant(); }
        }

        private static string SafeFormat(FileExportingEventArgs e)
        {
            try { return e.Format.ToString(); } catch { return ""; }
        }

        /// <summary>
        /// PDF の書き出しか。形式名（Revit 2022 以降の標準 PDF 書き出し）か、保存先の拡張子で判定する。
        /// </summary>
        internal static bool IsPdf(string format, string path)
        {
            if (!string.IsNullOrEmpty(format) && format.IndexOf("PDF", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            return !string.IsNullOrEmpty(path) && path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
        }

        private static string DescribeState(Document doc)
        {
            try
            {
                return $"(読み取り専用={doc.IsReadOnly}, リンク={doc.IsLinked}, トランザクション中={doc.IsModifiable}, ワークシェア={doc.IsWorkshared})";
            }
            catch (Exception ex) { return $"(取得失敗: {ex.Message})"; }
        }

        private static string SafeElementName(Element e)
        {
            try { return e.Name; } catch { return null; }
        }

        private static string SafeTypeName(Element e)
        {
            try
            {
                if (e is FamilySymbol fs) return $"{fs.FamilyName} : {fs.Name}";
                return e?.Name;
            }
            catch { return e?.Id?.ToString(); }
        }

        private static string SafeName(View v)
        {
            try { return v.Name; } catch { return v?.Id?.ToString() ?? "?"; }
        }

        /// <summary>警告ダイアログで出力を止めないよう、警告は黙って消す。</summary>
        private static void SetQuietFailureHandling(Transaction t)
        {
            var opt = t.GetFailureHandlingOptions();
            opt.SetFailuresPreprocessor(new SwallowWarnings());
            opt.SetClearAfterRollback(true);
            t.SetFailureHandlingOptions(opt);
        }

        private sealed class SwallowWarnings : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                foreach (var f in a.GetFailureMessages())
                {
                    if (f.GetSeverity() == FailureSeverity.Warning) a.DeleteWarning(f);
                }
                return FailureProcessingResult.Continue;
            }
        }
    }
}
