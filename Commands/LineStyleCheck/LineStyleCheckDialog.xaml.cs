using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Tools28.Commands.LineStyleCheck
{
    /// <summary>
    /// ListView 表示用データ
    /// </summary>
    public class ElementDisplayItem
    {
        public ElementId ElementId { get; set; }
        public string CategoryName { get; set; }
        public int OverriddenEdgeCount { get; set; }
        public string OverriddenStyleDisplay { get; set; }
    }

    public partial class LineStyleCheckDialog : Window
    {
        private readonly ExternalEvent _externalEvent;
        private readonly LineStyleCheckEventHandler _handler;
        private readonly UIDocument _uidoc;
        private List<OverriddenElementInfo> _overriddenElements;
        private bool _isClosingByUser = false;

        public LineStyleCheckDialog(
            ExternalEvent externalEvent,
            LineStyleCheckEventHandler handler,
            UIDocument uidoc,
            List<OverriddenElementInfo> overriddenElements)
        {
            InitializeComponent();

            _externalEvent = externalEvent;
            _handler = handler;
            _uidoc = uidoc;
            _overriddenElements = overriddenElements;

            _handler.Dialog = this;
            _handler.OverriddenElements = overriddenElements;

            // イベント登録
            _handler.OperationCompleted += OnOperationCompleted;
            _handler.RescanCompleted += OnRescanCompleted;

            // 線種リストを読み込み
            LoadLineStyles();

            // 検出結果を表示
            UpdateDisplay();
        }

        /// <summary>
        /// 線種コンボボックスを読み込む
        /// </summary>
        private void LoadLineStyles()
        {
            Document doc = _uidoc.Document;
            var lineStyles = LineStyleCheckEventHandler.GetAvailableLineStyles(doc);

            LineStyleComboBox.ItemsSource = lineStyles;

            if (lineStyles.Count > 0)
            {
                LineStyleComboBox.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// 検出結果の表示を更新する
        /// </summary>
        private void UpdateDisplay()
        {
            DetectedCountRun.Text = _overriddenElements.Count.ToString();

            // カテゴリ別の集計
            var categoryGroups = _overriddenElements
                .GroupBy(e => e.CategoryName)
                .Select(g => $"{g.Key}: {g.Count()}件")
                .ToList();

            DetailText.Text = string.Join("、", categoryGroups);

            // リストビューを更新
            var displayItems = _overriddenElements.Select(e => new ElementDisplayItem
            {
                ElementId = e.ElementId,
                CategoryName = e.CategoryName,
                OverriddenEdgeCount = e.OverriddenEdgeCount,
                OverriddenStyleDisplay = e.OverriddenStyleNames.Count > 0
                    ? string.Join(", ", e.OverriddenStyleNames)
                    : "-"
            }).ToList();

            ElementListView.ItemsSource = displayItems;

            ChangeAllButton.IsEnabled = _overriddenElements.Count > 0;
            PickChangeButton.IsEnabled = _overriddenElements.Count > 0;
        }

        /// <summary>
        /// 選択中の線種情報を取得する
        /// </summary>
        private LineStyleItem GetSelectedLineStyle()
        {
            return LineStyleComboBox.SelectedItem as LineStyleItem;
        }

        /// <summary>
        /// 一括変更ボタン
        /// </summary>
        private void ChangeAllButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = GetSelectedLineStyle();
            if (selected == null)
            {
                MessageBox.Show("変更先の線種を選択してください。", "線種変更確認",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _handler.RequestType = LineStyleRequestType.ChangeAll;
            _handler.SelectedGraphicsStyleId = selected.Id;
            _handler.IsResetToDefault = selected.IsResetToDefault;
            _handler.OverriddenElements = _overriddenElements;
            _externalEvent.Raise();
        }

        /// <summary>
        /// 選択変更ボタン
        /// </summary>
        private void PickChangeButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = GetSelectedLineStyle();
            if (selected == null)
            {
                MessageBox.Show("変更先の線種を選択してください。", "線種変更確認",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _handler.RequestType = LineStyleRequestType.PickAndChange;
            _handler.SelectedGraphicsStyleId = selected.Id;
            _handler.IsResetToDefault = selected.IsResetToDefault;
            _externalEvent.Raise();
        }

        /// <summary>
        /// 再スキャンボタン
        /// </summary>
        private void RescanButton_Click(object sender, RoutedEventArgs e)
        {
            _handler.RequestType = LineStyleRequestType.Rescan;
            _externalEvent.Raise();
        }

        /// <summary>
        /// 閉じるボタン
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _isClosingByUser = true;

            // ハイライトを解除してから閉じる
            _handler.RequestType = LineStyleRequestType.RemoveHighlight;
            _handler.OverriddenElements = _overriddenElements;
            _externalEvent.Raise();

            Close();
        }

        /// <summary>
        /// ウインドウの×ボタンで閉じた場合もハイライトを解除する
        /// </summary>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!_isClosingByUser)
            {
                _handler.RequestType = LineStyleRequestType.RemoveHighlight;
                _handler.OverriddenElements = _overriddenElements;
                _externalEvent.Raise();
            }

            _handler.OperationCompleted -= OnOperationCompleted;
            _handler.RescanCompleted -= OnRescanCompleted;
            _handler.Dialog = null;
        }

        /// <summary>
        /// 操作完了時のコールバック
        /// </summary>
        private void OnOperationCompleted(string resultMessage)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show(resultMessage, "線種変更確認",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }

        /// <summary>
        /// 再スキャン完了時のコールバック
        /// </summary>
        private void OnRescanCompleted(List<OverriddenElementInfo> newResults)
        {
            Dispatcher.Invoke(() =>
            {
                _overriddenElements = newResults;
                _handler.OverriddenElements = newResults;
                UpdateDisplay();

                if (newResults.Count == 0)
                {
                    MessageBox.Show(
                        "ラインワークで線種が変更された要素はありません。",
                        "線種変更確認",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            });
        }

        /// <summary>
        /// ダイアログを安全に表示する（外部イベントハンドラーから呼び出し用）
        /// </summary>
        public void SafeShow()
        {
            Dispatcher.Invoke(() =>
            {
                if (IsLoaded)
                {
                    Show();
                    Activate();
                }
            });
        }

        /// <summary>
        /// ダイアログを安全に非表示にする（外部イベントハンドラーから呼び出し用）
        /// </summary>
        public void SafeHide()
        {
            Dispatcher.Invoke(() =>
            {
                if (IsLoaded && IsVisible)
                {
                    Hide();
                }
            });
        }
    }
}
