using System;
using System.Collections.Generic;
using System.ComponentModel;   // ICollectionView（CollectionViewSource は System.Windows.Data）
using System.Linq;
using System.Windows;
using System.Windows.Data;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.DwgLayerTransfer.Models;
using Tools28.Commands.DwgLayerTransfer.Services;
using Tools28.Localization;

namespace Tools28.Commands.DwgLayerTransfer.Views
{
    /// <summary>
    /// 開いている別モデルから DWG レイヤの表示設定を取り込むダイアログ。
    /// 移行先は常にアクティブなモデル（このダイアログを開いたモデル）。
    ///
    /// 画面構成:
    ///   上部  … 移行元モデル / 移行先モデル / 移行の単位
    ///   左    … 移行元モデル データ情報（① ビュー → ② そのビューにある DWG）
    ///   右    … 移行先モデル データ情報（③ 反映するビュー（複数可） → ④ 反映先 DWG）
    /// </summary>
    public partial class DwgLayerTransferDialog : Window
    {
        private readonly Document _targetDoc;
        private readonly List<Document> _sourceDocs;
        private readonly DwgLayerScanner _scanner = new DwgLayerScanner();

        // --- 移行元 ---
        private Document _sourceDoc;
        private List<DwgDefinition> _sourceDwgs = new List<DwgDefinition>();
        private List<ViewEntry> _sourceViews = new List<ViewEntry>();
        private List<DwgItem> _sourceDwgItems = new List<DwgItem>();
        private ICollectionView _sourceViewsView;

        /// <summary>選択中の移行元ビュー×DWG のレイヤ設定（レイヤ名 -&gt; 設定、"" は DWG 本体）</summary>
        private Dictionary<string, LayerGraphicSetting> _sourceLayers
            = new Dictionary<string, LayerGraphicSetting>();

        // --- 移行先 ---
        private List<DwgDefinition> _targetDwgs = new List<DwgDefinition>();
        private List<TargetViewRow> _targetViewRows = new List<TargetViewRow>();
        private List<DwgItem> _targetDwgItems = new List<DwgItem>();
        private ICollectionView _targetViewsView;

        private bool _ready;   // 初期化完了フラグ（InitializeComponent 中のイベント発火を無視）

        public DwgLayerTransferDialog(Document targetDoc, List<Document> sourceDocs)
        {
            _targetDoc = targetDoc;
            _sourceDocs = sourceDocs ?? new List<Document>();

            InitializeComponent();
            ApplyLocalization();

            txtTarget.Text = SafeTitle(_targetDoc);

            foreach (var d in _sourceDocs)
                cmbSource.Items.Add(SafeTitle(d));
            if (cmbSource.Items.Count > 0) cmbSource.SelectedIndex = 0;

            _ready = true;
            ReloadAll();
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("DwgVg.Title");
            txtDescription.Text = Loc.S("DwgVg.Description");
            lblSource.Text = Loc.S("DwgVg.Source");
            lblTarget.Text = Loc.S("DwgVg.Target");
            lblMode.Text = Loc.S("DwgVg.Mode");
            rbTemplate.Content = Loc.S("DwgVg.Mode.Template");
            rbView.Content = Loc.S("DwgVg.Mode.View");
            lblModeHint.Text = Loc.S("DwgVg.Mode.Hint");

            lblSourceSection.Text = Loc.S("DwgVg.Section.Source");
            lblTargetSection.Text = Loc.S("DwgVg.Section.Target");

            // 移行元は常にビュー単位なので、①②の文言は固定
            lblSrcViewCaption.Text = Loc.S("DwgVg.Step1.View");
            lblSrcViewHint.Text = Loc.S("DwgVg.Step1.Hint");
            lblSrcDwgCaption.Text = Loc.S("DwgVg.Step2");
            lblTgtDwgCaption.Text = Loc.S("DwgVg.Step4");

            colTgtName.Header = Loc.S("DwgVg.Col.Name");
            colTgtStatus.Header = Loc.S("DwgVg.Col.Status");

            btnTgtSelectAll.Content = Loc.S("DwgVg.SelectAll");
            btnTgtDeselectAll.Content = Loc.S("DwgVg.DeselectAll");
            btnReload.Content = Loc.S("DwgVg.Btn.Reload");
            btnApply.Content = Loc.S("DwgVg.Btn.Apply");
            btnClose.Content = Loc.S("DwgVg.Btn.Close");

            UpdateModeDependentLabels();
        }

        /// <summary>反映先の単位（ビュー/ビューテンプレート）で表記が変わるラベルを更新する。</summary>
        private void UpdateModeDependentLabels()
        {
            lblTgtViewCaption.Text = TargetTemplateMode
                ? Loc.S("DwgVg.Step3.Template")
                : Loc.S("DwgVg.Step3");
        }

        private static string SafeTitle(Document d)
        {
            try { return d?.Title ?? ""; }
            catch { return ""; }
        }

        /// <summary>
        /// 反映先をビューテンプレート単位にするか（false ならビュー単位）。
        /// 移行元は常にビュー単位で読み取るため、この選択は移行先にだけ効く。
        /// ビューから読むことで、ビューテンプレート由来の設定と
        /// ビュー個別に設定された V/G のどちらでも「今そのビューで効いている値」を取得できる。
        /// </summary>
        private bool TargetTemplateMode => rbTemplate.IsChecked == true;

        // ===== 一覧の構築 =====

        /// <summary>移行元モデル・移行の単位が変わったときに全体を組み直す。</summary>
        private void ReloadAll()
        {
            if (!_ready) return;

            int i = cmbSource.SelectedIndex;
            _sourceDoc = (i >= 0 && i < _sourceDocs.Count) ? _sourceDocs[i] : null;

            UpdateModeDependentLabels();

            if (_sourceDoc == null)
            {
                ClearAll();
                return;
            }

            // 走査中は Revit 本体の操作を無効化する（using を抜けると必ず復帰）
            using (this.BlockRevitInput())
            {
                _sourceDwgs = _scanner.EnumerateDwgs(_sourceDoc);
                _targetDwgs = _scanner.EnumerateDwgs(_targetDoc);

                // 移行元は常に「ビュー」。テンプレート側の V/G で「読み込み」が
                // 含められていない場合でも、ビューから読めば実際に効いている値が取れる
                _sourceViews = _scanner.EnumerateViews(_sourceDoc, templates: false, checkTemplateControl: false);
                var targetViews = _scanner.EnumerateViews(_targetDoc, TargetTemplateMode, checkTemplateControl: true);

                // --- 左: ビュー一覧 ---
                _sourceViewsView = CollectionViewSource.GetDefaultView(_sourceViews);
                _sourceViewsView.Filter = o =>
                    o is ViewEntry v && MatchesSearch(v.Name, SrcViewSearch?.Text);
                SrcViewList.ItemsSource = _sourceViewsView;

                // --- 右: ビュー一覧 ---
                _targetViewRows = targetViews.Select(v => new TargetViewRow { Entry = v }).ToList();
                foreach (var r in _targetViewRows) r.PropertyChanged += TargetViewRow_PropertyChanged;

                _targetViewsView = CollectionViewSource.GetDefaultView(_targetViewRows);
                _targetViewsView.Filter = o =>
                    o is TargetViewRow r && MatchesSearch(r.Name, TgtViewSearch?.Text);
                TgtViewGrid.ItemsSource = _targetViewsView;

                // --- 右: DWG 一覧（移行先は全 DWG を出す。絞り込むと選べなくなるため）---
                _targetDwgItems = _targetDwgs.Select(d => new DwgItem { Dwg = d }).ToList();
                TgtDwgList.ItemsSource = _targetDwgItems;

                // 左のビューは先頭を自動選択（→ 連鎖して DWG 一覧まで埋まる）
                SrcViewList.SelectedIndex = _sourceViews.Count > 0 ? 0 : -1;
                if (SrcViewList.SelectedIndex < 0) OnSourceViewChanged();
            }

            UpdateSummary();
        }

        private void ClearAll()
        {
            _sourceDwgs = new List<DwgDefinition>();
            _sourceViews = new List<ViewEntry>();
            _sourceDwgItems = new List<DwgItem>();
            _targetDwgs = new List<DwgDefinition>();
            _targetViewRows = new List<TargetViewRow>();
            _targetDwgItems = new List<DwgItem>();
            _sourceLayers = new Dictionary<string, LayerGraphicSetting>();

            SrcViewList.ItemsSource = null;
            SrcDwgList.ItemsSource = null;
            TgtViewGrid.ItemsSource = null;
            TgtDwgList.ItemsSource = null;
            UpdateSummary();
        }

        private static bool MatchesSearch(string name, string query)
        {
            string q = query?.Trim();
            if (string.IsNullOrEmpty(q)) return true;
            return name != null && name.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        /// <summary>
        /// ① 移行元ビューが変わったとき: そのビューにある DWG の一覧を作り直す。
        /// 設定件数もここで数えるが、対象は「選択中の1ビュー × その DWG」だけなので軽い。
        /// </summary>
        private void OnSourceViewChanged()
        {
            var entry = SrcViewList.SelectedItem as ViewEntry;

            if (entry == null || _sourceDoc == null)
            {
                _sourceDwgItems = new List<DwgItem>();
                SrcDwgList.ItemsSource = null;
                _sourceLayers = new Dictionary<string, LayerGraphicSetting>();
                SyncTargetViewSelection(null);
                UpdateSummary();
                return;
            }

            var visible = DwgLayerScanner.FilterForView(_sourceDwgs, entry.Id, entry.IsTemplate);

            _sourceDwgItems = new List<DwgItem>();
            foreach (var d in visible)
            {
                var settings = _scanner.ReadSettings(_sourceDoc, entry.Id, d);
                _sourceDwgItems.Add(new DwgItem
                {
                    Dwg = d,
                    SettingCount = DwgLayerScanner.CountConfigured(settings.Values)
                });
            }

            // 設定があるものを上に出して見つけやすくする
            _sourceDwgItems = _sourceDwgItems
                .OrderByDescending(x => x.SettingCount > 0)
                .ThenBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            SrcDwgList.ItemsSource = _sourceDwgItems;

            // 同名の移行先ビューを自動でチェック
            SyncTargetViewSelection(entry.Name);

            // DWG は設定があるものを優先して自動選択
            SrcDwgList.SelectedItem = _sourceDwgItems.FirstOrDefault(x => x.SettingCount > 0)
                                      ?? _sourceDwgItems.FirstOrDefault();
            if (SrcDwgList.SelectedItem == null) OnSourceDwgChanged();
        }

        /// <summary>② 移行元 DWG が変わったとき: レイヤ設定を読み、移行先 DWG の照合を更新する。</summary>
        private void OnSourceDwgChanged()
        {
            var item = SrcDwgList.SelectedItem as DwgItem;
            var entry = SrcViewList.SelectedItem as ViewEntry;

            if (item?.Dwg == null || entry == null || _sourceDoc == null)
            {
                _sourceLayers = new Dictionary<string, LayerGraphicSetting>();
                foreach (var t in _targetDwgItems) t.MatchedLayerCount = -1;
                UpdateSummary();
                return;
            }

            _sourceLayers = _scanner.ReadSettings(_sourceDoc, entry.Id, item.Dwg);

            // 移行先 DWG それぞれについて、レイヤ名がいくつ一致するかを出す
            foreach (var t in _targetDwgItems)
                t.MatchedLayerCount = t.Dwg == null
                    ? 0
                    : item.Dwg.Layers.Keys.Count(t.Dwg.Layers.ContainsKey);

            // 同名の移行先 DWG を自動選択、無ければ一致レイヤが最も多いもの
            var same = _targetDwgItems.FirstOrDefault(
                t => string.Equals(t.Name, item.Name, StringComparison.CurrentCultureIgnoreCase));
            TgtDwgList.SelectedItem = same
                ?? _targetDwgItems.OrderByDescending(t => t.MatchedLayerCount).FirstOrDefault(t => t.MatchedLayerCount > 0);

            UpdateSummary();
        }

        /// <summary>③ 移行元ビューと同名の移行先ビューだけをチェック状態にする。</summary>
        private void SyncTargetViewSelection(string sourceViewName)
        {
            foreach (var r in _targetViewRows)
                r.IsSelected = sourceViewName != null
                            && string.Equals(r.Name, sourceViewName, StringComparison.CurrentCultureIgnoreCase);

            try { TgtViewGrid.Items.Refresh(); } catch { }
        }

        private void UpdateSummary()
        {
            var srcView = SrcViewList.SelectedItem as ViewEntry;
            var srcDwg = SrcDwgList.SelectedItem as DwgItem;
            var tgtDwg = TgtDwgList.SelectedItem as DwgItem;
            int tgtViewCount = _targetViewRows.Count(r => r.IsSelected);

            int configured = DwgLayerScanner.CountConfigured(_sourceLayers.Values);

            txtSummary.Text = string.Format(
                Loc.S("DwgVg.Summary"),
                srcView?.Name ?? "-",
                srcDwg?.Name ?? "-",
                configured,
                tgtViewCount,
                tgtDwg?.Name ?? "-");

            btnApply.IsEnabled = srcView != null && srcDwg != null && tgtDwg != null
                                 && tgtViewCount > 0 && _sourceLayers.Count > 0;
        }

        // ===== イベント =====

        private void Source_Changed(object sender, RoutedEventArgs e) => ReloadAll();

        private void Mode_Changed(object sender, RoutedEventArgs e) => ReloadAll();

        private void Reload_Click(object sender, RoutedEventArgs e) => ReloadAll();

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void SrcViewSearch_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            try { _sourceViewsView?.Refresh(); } catch { }
        }

        private void TgtViewSearch_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            try { _targetViewsView?.Refresh(); } catch { }
        }

        private void SrcView_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            OnSourceViewChanged();
        }

        private void SrcDwg_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            OnSourceDwgChanged();
        }

        private void TgtDwg_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;
            UpdateSummary();
        }

        private void TargetViewRow_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(TargetViewRow.IsSelected)) UpdateSummary();
        }

        private void TgtSelectAll_Click(object sender, RoutedEventArgs e) => SetTargetSelection(true);

        private void TgtDeselectAll_Click(object sender, RoutedEventArgs e) => SetTargetSelection(false);

        /// <summary>検索で表示中の行だけを対象にチェックを付け外しする。</summary>
        private void SetTargetSelection(bool selected)
        {
            if (_targetViewsView == null) return;
            foreach (var o in _targetViewsView.Cast<object>())
                if (o is TargetViewRow r) r.IsSelected = selected;

            try { TgtViewGrid.Items.Refresh(); } catch { }
            UpdateSummary();
        }

        // ===== 適用 =====

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            var srcView = SrcViewList.SelectedItem as ViewEntry;
            var srcDwg = SrcDwgList.SelectedItem as DwgItem;
            var tgtDwg = TgtDwgList.SelectedItem as DwgItem;

            if (srcView == null) { Warn("DwgVg.Warn.NoSourceView"); return; }
            if (srcDwg?.Dwg == null) { Warn("DwgVg.Warn.NoSourceDwg"); return; }
            if (tgtDwg?.Dwg == null) { Warn("DwgVg.Warn.NoTargetDwg"); return; }

            var targetViews = _targetViewRows
                .Where(r => r.IsSelected && r.IsApplicable)
                .Select(r => r.Entry)
                .ToList();
            if (targetViews.Count == 0) { Warn("DwgVg.Warn.NoTargetView"); return; }

            if (_sourceLayers.Count == 0) { Warn("DwgVg.Warn.NoLayers"); return; }

            var confirm = new TaskDialog(Loc.S("DwgVg.Confirm.Title"))
            {
                MainInstruction = string.Format(Loc.S("DwgVg.Confirm.Main"),
                    srcView.Name, srcDwg.Name, targetViews.Count, tgtDwg.Name),
                MainContent = Loc.S("DwgVg.Confirm.Content"),
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                DefaultButton = TaskDialogResult.No
            };
            if (confirm.Show() != TaskDialogResult.Yes)
            {
                this.BringToFrontDeferred();
                return;
            }

            TransferResult result;
            try
            {
                using (this.BlockRevitInput())
                {
                    result = new DwgLayerApplier().Apply(
                        _targetDoc, _sourceLayers, targetViews, tgtDwg.Dwg);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show(Loc.S("Common.Error"),
                    string.Format(Loc.S("DwgVg.Result.Error"), ex.Message));
                this.BringToFrontDeferred();
                return;
            }

            ShowResult(result);
            this.BringToFrontDeferred();
        }

        private void Warn(string key)
        {
            TaskDialog.Show(Loc.S("DwgVg.Title"), Loc.S(key));
            this.BringToFrontDeferred();
        }

        private static void ShowResult(TransferResult r)
        {
            var content = new System.Text.StringBuilder();

            if (r.MissingLayers.Count > 0)
            {
                content.AppendLine(string.Format(Loc.S("DwgVg.Result.MissingLayers"), r.MissingLayers.Count));
                content.AppendLine(BuildPreview(r.MissingLayers));
                content.AppendLine();
            }

            if (r.MissingLinePatterns.Count > 0)
            {
                content.AppendLine(string.Format(Loc.S("DwgVg.Result.MissingPatterns"), r.MissingLinePatterns.Count));
                content.AppendLine(BuildPreview(r.MissingLinePatterns));
            }

            var dlg = new TaskDialog(Loc.S("DwgVg.Result.Title"))
            {
                MainInstruction = string.Format(Loc.S("DwgVg.Result.Msg"), r.ViewCount, r.LayerCount),
                MainContent = content.ToString().TrimEnd()
            };
            dlg.Show();
        }

        /// <summary>未解決一覧は先頭 10 件までを表示し、残りは件数だけ添える。</summary>
        private static string BuildPreview(List<string> items)
        {
            const int max = 10;
            var head = items.Take(max).Select(s => "  ・" + s);
            string text = string.Join("\n", head);
            if (items.Count > max)
                text += "\n  " + string.Format(Loc.S("DwgVg.Result.More"), items.Count - max);
            return text;
        }
    }
}
