using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools28.Commands.FilterManagement.Models;
using Tools28.Commands.FilterManagement.Services;
using Tools28.Localization;

namespace Tools28.Commands.FilterManagement.Views
{
    public partial class FilterManagementDialog : Window
    {
        private readonly Document _doc;
        private readonly FilterScanner _scanner = new FilterScanner();

        private List<FilterRow> _rows = new List<FilterRow>();
        private ICollectionView _view;
        private bool _ready;   // 初期化完了フラグ（InitializeComponent 中のイベント発火を無視）

        public FilterManagementDialog(Document doc)
        {
            _doc = doc;
            InitializeComponent();
            ApplyLocalization();
            LoadRows();
            _ready = true;
        }

        private void ApplyLocalization()
        {
            Title = Loc.S("FilterMgmt.Title");
            txtDescription.Text = Loc.S("FilterMgmt.Description");
            lblSearch.Text = Loc.S("FilterMgmt.Search");
            chkUnusedOnly.Content = Loc.S("FilterMgmt.UnusedOnly");
            lblKindFilter.Text = Loc.S("FilterMgmt.Filter.Kind");
            rbKindAll.Content = Loc.S("FilterMgmt.Filter.All");
            rbKindRule.Content = Loc.S("FilterMgmt.Kind.Rule");
            rbKindSelection.Content = Loc.S("FilterMgmt.Kind.Selection");

            FilterGrid.Columns[0].Header = Loc.S("FilterMgmt.Col.Select");
            colName.Header = Loc.S("FilterMgmt.Col.Name");
            colKind.Header = Loc.S("FilterMgmt.Col.Kind");
            colCategories.Header = Loc.S("FilterMgmt.Col.Categories");
            colViewCount.Header = Loc.S("FilterMgmt.Col.ViewCount");
            colUsage.Header = Loc.S("FilterMgmt.Col.Usage");

            btnSelectUnused.Content = Loc.S("FilterMgmt.SelectUnused");
            btnSelectAllVisible.Content = Loc.S("FilterMgmt.SelectAll");
            btnDeselectAllVisible.Content = Loc.S("FilterMgmt.DeselectAll");
            btnReload.Content = Loc.S("FilterMgmt.Btn.Reload");
            btnDelete.Content = Loc.S("FilterMgmt.Btn.Delete");
            btnClose.Content = Loc.S("FilterMgmt.Btn.Close");
        }

        private void LoadRows()
        {
            // 全ビュー走査中は Revit 本体の操作を無効化する（using を抜けると必ず復帰）
            using (this.BlockRevitInput())
            {
                _rows = _scanner.EnumerateFilters(_doc);
                _view = CollectionViewSource.GetDefaultView(_rows);
                _view.Filter = RowFilter;
                FilterGrid.ItemsSource = _view;
                UpdateCount();
            }
        }

        private bool RowFilter(object o)
        {
            if (!(o is FilterRow r)) return false;

            string q = SearchBox?.Text?.Trim();
            if (!string.IsNullOrEmpty(q) &&
                (r.Name == null || r.Name.IndexOf(q, StringComparison.CurrentCultureIgnoreCase) < 0))
                return false;

            if (chkUnusedOnly?.IsChecked == true && r.IsUsed)
                return false;

            if (rbKindRule?.IsChecked == true && r.Kind != FilterKind.Rule) return false;
            if (rbKindSelection?.IsChecked == true && r.Kind != FilterKind.Selection) return false;

            return true;
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (!_ready) return;   // 初期化中のイベントは無視
            _view?.Refresh();
            UpdateCount();
        }

        private void UpdateCount()
        {
            int total = _rows.Count;
            int shown = _view?.Cast<object>().Count() ?? total;
            int unused = _rows.Count(r => !r.IsUsed);
            int selected = _rows.Count(r => r.IsSelected);
            txtCount.Text = string.Format(Loc.S("FilterMgmt.Count.Summary"), total, shown, unused, selected);
        }

        // ===== 名前のインライン変更 =====
        // セル編集確定のイベント内で Revit トランザクションや TaskDialog を実行すると
        // DataGrid の編集状態と再入し不安定になるため、実際のリネームは編集確定後に
        // Dispatcher（Background 優先度）で遅延実行する。
        private void FilterGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit) return;
            if (e.Column != colName) return;
            if (!(e.Row.Item is FilterRow row)) return;
            if (!(e.EditingElement is System.Windows.Controls.TextBox tb)) return;

            string oldName = row.Name;              // 変更前の確定名（バインドはまだ未反映）
            string newName = tb.Text?.Trim() ?? "";
            if (newName == oldName) return;

            Dispatcher.BeginInvoke(
                new Action(() => TryRename(row, oldName, newName)),
                System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>1件のフィルタ名変更を検証してトランザクションで適用する（失敗時は元へ戻す）。</summary>
        private void TryRename(FilterRow row, string oldName, string newName)
        {
            // 空名は不可
            if (string.IsNullOrEmpty(newName))
            {
                TaskDialog.Show(Loc.S("FilterMgmt.Title"), Loc.S("FilterMgmt.Rename.Empty"));
                row.Name = oldName;
                this.BringToFrontDeferred();
                return;
            }

            // 同名重複（大文字小文字無視）を事前チェック
            if (_rows.Any(r => r != row &&
                               string.Equals(r.Name, newName, StringComparison.CurrentCultureIgnoreCase)))
            {
                TaskDialog.Show(Loc.S("FilterMgmt.Title"),
                    string.Format(Loc.S("FilterMgmt.Rename.Duplicate"), newName));
                row.Name = oldName;
                this.BringToFrontDeferred();
                return;
            }

            var fe = _doc.GetElement(row.Id);
            if (fe == null) { row.Name = oldName; return; }

            try
            {
                using (var t = new Transaction(_doc, Loc.S("FilterMgmt.Txn.Rename")))
                {
                    t.Start();
                    fe.Name = newName;
                    t.Commit();
                }
                row.Name = newName;   // 念のため確定
                UpdateCount();
            }
            catch (Exception ex)
            {
                TaskDialog.Show(Loc.S("Common.Error"),
                    string.Format(Loc.S("FilterMgmt.Rename.Error"), ex.Message));
                row.Name = oldName;   // 元の名前へ戻す
                this.BringToFrontDeferred();
            }
        }

        // ===== 選択操作 =====
        private void SelectUnused_Click(object sender, RoutedEventArgs e)
        {
            foreach (var r in _rows)
                r.IsSelected = !r.IsUsed;
            _view?.Refresh();
            UpdateCount();
        }

        private void SelectAllVisible_Click(object sender, RoutedEventArgs e) => SetSelectionForVisible(true);
        private void DeselectAllVisible_Click(object sender, RoutedEventArgs e) => SetSelectionForVisible(false);

        private void SetSelectionForVisible(bool selected)
        {
            if (_view == null) return;
            foreach (var o in _view.Cast<object>())
                if (o is FilterRow r) r.IsSelected = selected;
            UpdateCount();
        }

        // ===== 削除 =====
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            CommitPendingEdit();

            var selected = _rows.Where(r => r.IsSelected).ToList();
            if (selected.Count == 0)
            {
                TaskDialog.Show(Loc.S("FilterMgmt.Title"), Loc.S("FilterMgmt.NoSelection.Msg"));
                this.BringToFrontDeferred();
                return;
            }

            int usedCount = selected.Count(r => r.IsUsed);
            var confirm = new TaskDialog(Loc.S("FilterMgmt.Confirm.Title"))
            {
                MainInstruction = string.Format(Loc.S("FilterMgmt.Confirm.Main"), selected.Count),
                MainContent = usedCount > 0
                    ? string.Format(Loc.S("FilterMgmt.Confirm.ContentUsed"), usedCount)
                    : Loc.S("FilterMgmt.Confirm.Content"),
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                DefaultButton = TaskDialogResult.No
            };
            if (confirm.Show() != TaskDialogResult.Yes)
                return;

            var ids = selected.Select(r => r.Id)
                              .Where(id => id != null && id != ElementId.InvalidElementId)
                              .ToList();
            int ok = 0, fail = 0;
            List<string> warnings = null;   // 削除時に Revit が出した警告（記録して後で通知）

            try
            {
                using (var t = new Transaction(_doc, Loc.S("FilterMgmt.Txn.Delete")))
                {
                    t.Start();

                    // 削除に伴う Revit の警告ダイアログを自動で閉じる
                    var fho = t.GetFailureHandlingOptions();
                    fho = fho.SetForcedModalHandling(false);
                    fho = fho.SetClearAfterRollback(true);
                    var swallower = new WarningSwallower();
                    warnings = swallower.Messages;
                    fho = fho.SetFailuresPreprocessor(swallower);
                    t.SetFailureHandlingOptions(fho);

                    ICollection<ElementId> deleted;
                    try
                    {
                        deleted = _doc.Delete(ids);
                    }
                    catch
                    {
                        // 一括削除に失敗した場合は個別に試行
                        deleted = new List<ElementId>();
                        foreach (var id in ids)
                        {
                            try
                            {
                                var r = _doc.Delete(id);
                                if (r != null)
                                    foreach (var did in r) deleted.Add(did);
                            }
                            catch { }
                        }
                    }
                    t.Commit();

                    var deletedSet = new HashSet<ElementId>(deleted ?? new List<ElementId>());
                    ok = ids.Count(id => deletedSet.Contains(id));
                    fail = ids.Count - ok;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show(Loc.S("Common.Error"),
                    string.Format(Loc.S("FilterMgmt.Result.Error"), ex.Message));
                this.BringToFrontDeferred();
                return;
            }

            string resultMsg = string.Format(Loc.S("FilterMgmt.Result.Msg"), ok, fail);
            if (warnings != null && warnings.Count > 0)
            {
                resultMsg += "\n\n" + Loc.S("Common.RevitWarnings") + "\n"
                           + string.Join("\n", warnings.Take(10).Select(w => "・" + w));
                if (warnings.Count > 10) resultMsg += "\n…";
            }
            TaskDialog.Show(Loc.S("FilterMgmt.Result.Title"), resultMsg);

            LoadRows();   // 一覧を再構築
            this.BringToFrontDeferred();   // Revit 本体の背面に隠れるのを防ぐ
        }

        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            CommitPendingEdit();
            LoadRows();
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        /// <summary>編集中のセルがあれば確定させる（ボタン操作前に呼ぶ）。</summary>
        private void CommitPendingEdit()
        {
            try
            {
                FilterGrid.CommitEdit(DataGridEditingUnit.Cell, true);
                FilterGrid.CommitEdit(DataGridEditingUnit.Row, true);
            }
            catch { }
        }

        /// <summary>
        /// 削除トランザクション中の警告をモーダル表示せずに処理する。
        /// ただし警告文は捨てずに記録し、削除後に結果ダイアログでまとめて通知する
        /// （「要素が削除されます」等の重要な警告を見落とさないため）。
        /// </summary>
        private class WarningSwallower : IFailuresPreprocessor
        {
            public List<string> Messages { get; } = new List<string>();

            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                try
                {
                    foreach (var f in a.GetFailureMessages())
                    {
                        var text = f?.GetDescriptionText();
                        if (!string.IsNullOrWhiteSpace(text) && !Messages.Contains(text))
                            Messages.Add(text);
                    }
                }
                catch { }

                a.DeleteAllWarnings();
                return FailureProcessingResult.Continue;
            }
        }
    }
}
