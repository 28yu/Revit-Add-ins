using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
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
    /// </summary>
    public partial class DwgLayerTransferDialog : Window
    {
        private readonly Document _targetDoc;
        private readonly List<Document> _sourceDocs;
        private readonly DwgLayerScanner _scanner = new DwgLayerScanner();

        private List<DwgDefinition> _targetDwgs = new List<DwgDefinition>();
        private List<DwgPairRow> _dwgRows = new List<DwgPairRow>();
        private List<ViewPairRow> _viewRows = new List<ViewPairRow>();

        /// <summary>移行先のビュー名 -&gt; ビュー情報（現在のモードで列挙したもの）</summary>
        private Dictionary<string, ViewEntry> _targetViewsByName
            = new Dictionary<string, ViewEntry>(StringComparer.CurrentCultureIgnoreCase);

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
            Rebuild();
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

            lblDwgMap.Text = Loc.S("DwgVg.DwgMap");
            colDwgSource.Header = Loc.S("DwgVg.Col.DwgSource");
            colDwgTarget.Header = Loc.S("DwgVg.Col.DwgTarget");
            colDwgStatus.Header = Loc.S("DwgVg.Col.DwgStatus");

            lblViewMap.Text = Loc.S("DwgVg.ViewMap");
            colViewSelect.Header = Loc.S("DwgVg.Col.Select");
            colViewSource.Header = Loc.S("DwgVg.Col.ViewSource");
            colViewCount.Header = Loc.S("DwgVg.Col.SettingCount");
            colViewTarget.Header = Loc.S("DwgVg.Col.ViewTarget");
            colViewStatus.Header = Loc.S("DwgVg.Col.Status");

            btnSelectWithSettings.Content = Loc.S("DwgVg.SelectWithSettings");
            btnSelectWithSettings.ToolTip = Loc.S("DwgVg.SelectWithSettings.Tip");
            btnSelectAll.Content = Loc.S("DwgVg.SelectAll");
            btnDeselectAll.Content = Loc.S("DwgVg.DeselectAll");
            btnReload.Content = Loc.S("DwgVg.Btn.Reload");
            btnApply.Content = Loc.S("DwgVg.Btn.Apply");
            btnClose.Content = Loc.S("DwgVg.Btn.Close");
        }

        private static string SafeTitle(Document d)
        {
            try { return d?.Title ?? ""; }
            catch { return ""; }
        }

        private Document SelectedSourceDoc
        {
            get
            {
                int i = cmbSource.SelectedIndex;
                return (i >= 0 && i < _sourceDocs.Count) ? _sourceDocs[i] : null;
            }
        }

        /// <summary>ビューテンプレート単位か（false ならビュー単位）</summary>
        private bool TemplateMode => rbTemplate.IsChecked == true;

        // ===== 一覧の構築 =====

        private void Rebuild()
        {
            if (!_ready) return;

            var src = SelectedSourceDoc;
            if (src == null)
            {
                _dwgRows = new List<DwgPairRow>();
                _viewRows = new List<ViewPairRow>();
                DwgGrid.ItemsSource = null;
                ViewGrid.ItemsSource = null;
                btnApply.IsEnabled = false;
                UpdateCount();
                return;
            }

            // 走査中は Revit 本体の操作を無効化する（using を抜けると必ず復帰）
            using (this.BlockRevitInput())
            {
                _targetDwgs = _scanner.EnumerateDwgs(_targetDoc);
                var sourceDwgs = _scanner.EnumerateDwgs(src);

                BuildDwgRows(sourceDwgs);
                BuildViewRows(src, sourceDwgs);
            }

            btnApply.IsEnabled = _viewRows.Any(r => r.IsApplicable);
            UpdateCount();
        }

        private void BuildDwgRows(List<DwgDefinition> sourceDwgs)
        {
            var candidates = new List<string> { DwgPairRow.NoMatch };
            candidates.AddRange(_targetDwgs.Select(d => d.Name));

            _dwgRows = new List<DwgPairRow>();
            foreach (var s in sourceDwgs)
            {
                var row = new DwgPairRow
                {
                    Source = s,
                    Candidates = candidates
                };

                // 同名の DWG を自動で対応付ける
                var match = _targetDwgs.FirstOrDefault(
                    d => string.Equals(d.Name, s.Name, StringComparison.CurrentCultureIgnoreCase));
                row.SelectedTarget = match?.Name ?? DwgPairRow.NoMatch;

                UpdateMatchedLayerCount(row);
                row.PropertyChanged += DwgRow_PropertyChanged;
                _dwgRows.Add(row);
            }

            DwgGrid.ItemsSource = _dwgRows;
        }

        private void DwgRow_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(DwgPairRow.SelectedTarget)) return;
            if (!(sender is DwgPairRow row)) return;

            UpdateMatchedLayerCount(row);
            // StatusText は MatchedLayerCount を参照するため、更新後に再通知させる。
            // ComboBox の選択確定中に Items.Refresh() を呼ぶと DataGrid の編集トランザクションと
            // 衝突して例外になるため、Background 優先度で遅延実行する。
            RefreshGridDeferred(DwgGrid);
        }

        /// <summary>DataGrid の再描画を編集確定後まで遅らせる。</summary>
        private void RefreshGridDeferred(System.Windows.Controls.DataGrid grid)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    grid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
                    grid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
                    grid.Items.Refresh();
                }
                catch { }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>移行元レイヤ名のうち、選択中の移行先 DWG に同名が存在する件数を数える。</summary>
        private void UpdateMatchedLayerCount(DwgPairRow row)
        {
            if (!row.IsResolved) { row.MatchedLayerCount = 0; return; }

            var target = _targetDwgs.FirstOrDefault(
                d => string.Equals(d.Name, row.SelectedTarget, StringComparison.CurrentCultureIgnoreCase));
            if (target == null) { row.MatchedLayerCount = 0; return; }

            row.MatchedLayerCount = row.Source.Layers.Keys.Count(target.Layers.ContainsKey);
        }

        private void BuildViewRows(Document src, List<DwgDefinition> sourceDwgs)
        {
            bool templates = TemplateMode;

            // 移行先はテンプレート制御の有無を判定する（ビュー単位＝テンプレートに奪われていないか、
            // テンプレート単位＝そのテンプレートが読み込みカテゴリを制御しているか）
            var targetViews = _scanner.EnumerateViews(_targetDoc, templates, checkTemplateControl: true);

            _targetViewsByName = new Dictionary<string, ViewEntry>(StringComparer.CurrentCultureIgnoreCase);
            foreach (var v in targetViews) _targetViewsByName[v.Name] = v;

            var candidates = new List<string> { ViewPairRow.NoMatch };
            candidates.AddRange(targetViews.Select(v => v.Name));

            var sourceViews = _scanner.EnumerateViews(src, templates, checkTemplateControl: false);
            var snapshots = _scanner.ReadSettings(src, sourceViews, sourceDwgs);

            _viewRows = new List<ViewPairRow>();
            foreach (var snap in snapshots)
            {
                var row = new ViewPairRow
                {
                    Snapshot = snap,
                    Candidates = candidates
                };

                if (_targetViewsByName.TryGetValue(snap.View.Name, out var match))
                {
                    row.BlockReason = match.BlockReason;
                    row.SelectedTarget = match.Name;
                }
                else
                {
                    row.SelectedTarget = ViewPairRow.NoMatch;
                }

                // 設定が1件以上あり、そのまま適用できる行を既定でチェック
                row.IsSelected = row.IsApplicable && row.SettingCount > 0;

                row.PropertyChanged += ViewRow_PropertyChanged;
                _viewRows.Add(row);
            }

            // 設定があるものを上に、その中は名前昇順
            _viewRows = _viewRows
                .OrderByDescending(r => r.SettingCount > 0)
                .ThenBy(r => r.SourceName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            ViewGrid.ItemsSource = _viewRows;
        }

        private void ViewRow_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (!(sender is ViewPairRow row)) return;

            if (e.PropertyName == nameof(ViewPairRow.SelectedTarget))
            {
                // 対応先を変えたらテンプレート制御の判定も付け替える
                string reason = null;
                if (row.IsResolved && _targetViewsByName.TryGetValue(row.SelectedTarget, out var entry))
                    reason = entry.BlockReason;
                row.BlockReason = reason;

                RefreshGridDeferred(ViewGrid);
            }

            if (e.PropertyName == nameof(ViewPairRow.SelectedTarget) ||
                e.PropertyName == nameof(ViewPairRow.IsSelected))
            {
                UpdateCount();
            }
        }

        private void UpdateCount()
        {
            int total = _viewRows.Count;
            int selected = _viewRows.Count(r => r.IsSelected);
            int withSettings = _viewRows.Count(r => r.SettingCount > 0);
            txtCount.Text = string.Format(Loc.S("DwgVg.Count.Summary"), total, selected, withSettings);
        }

        // ===== イベント =====

        private void Source_Changed(object sender, RoutedEventArgs e) => Rebuild();

        private void Mode_Changed(object sender, RoutedEventArgs e) => Rebuild();

        private void Reload_Click(object sender, RoutedEventArgs e) => Rebuild();

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void SelectWithSettings_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _viewRows) r.IsSelected = r.IsApplicable && r.SettingCount > 0;
            ViewGrid.Items.Refresh();
            UpdateCount();
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _viewRows) r.IsSelected = r.IsApplicable;
            ViewGrid.Items.Refresh();
            UpdateCount();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _viewRows) r.IsSelected = false;
            ViewGrid.Items.Refresh();
            UpdateCount();
        }

        // ===== 適用 =====

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            var targets = _viewRows.Where(r => r.IsSelected && r.IsApplicable).ToList();
            if (targets.Count == 0)
            {
                TaskDialog.Show(Loc.S("DwgVg.Title"), Loc.S("DwgVg.NoSelection"));
                this.BringToFrontDeferred();
                return;
            }

            // 同名の DWG カテゴリが複数存在しうるため ToDictionary は使わず、後勝ちで詰める
            var dwgMap = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            foreach (var r in _dwgRows.Where(r => r.IsResolved))
                dwgMap[r.Source.Name] = r.SelectedTarget;

            if (dwgMap.Count == 0)
            {
                TaskDialog.Show(Loc.S("DwgVg.Title"), Loc.S("DwgVg.NoDwgMapping"));
                this.BringToFrontDeferred();
                return;
            }

            var confirm = new TaskDialog(Loc.S("DwgVg.Confirm.Title"))
            {
                MainInstruction = string.Format(Loc.S("DwgVg.Confirm.Main"), targets.Count, dwgMap.Count),
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
                        _targetDoc, targets, _targetViewsByName, dwgMap, _targetDwgs);
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
