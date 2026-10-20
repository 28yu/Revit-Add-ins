using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ComboBox = System.Windows.Controls.ComboBox;
using Grid = System.Windows.Controls.Grid;
using Tools28.Localization;

namespace Tools28.Commands.RealtimeRender
{
    /// <summary>
    /// フェーズの「グラフィックの上書き」のマテリアル（既存・解体などの状態ごと）。
    /// </summary>
    internal class PhaseOverrides
    {
        public ElementId ViewPhaseId = ElementId.InvalidElementId;
        /// <summary>状態 → 使うマテリアル（このビューで「上書き」表示になっている状態だけ）。</summary>
        public Dictionary<ElementOnPhaseStatus, ElementId> Materials = new Dictionary<ElementOnPhaseStatus, ElementId>();
    }

    /// <summary>
    /// 3D ビューのフェーズ フィルタで「上書き」表示になっている要素を、フェーズのグラフィックの上書きのマテリアルで描くための準備。
    ///
    /// Revit API ではフェーズの「グラフィックの上書き」のマテリアルを読めないので、
    /// ① 既定の名前（「フェーズ - 既存」「Phase - Exist」など）で探し、② 利用者に選んでもらい、その選択を PC に覚えておく。
    /// 覚える場所は %APPDATA%\Tools28\RealtimeRender\phase-materials.txt（モデルごと。共有モデルでも書き込み権限が要らないように）。
    /// </summary>
    internal static class PhaseMaterials
    {
        private static readonly ElementOnPhaseStatus[] Statuses =
        {
            ElementOnPhaseStatus.Existing, ElementOnPhaseStatus.Demolished, ElementOnPhaseStatus.New, ElementOnPhaseStatus.Temporary,
        };

        // 状態ごとの名前の手がかり（日本語・英語・中国語のテンプレートの名前）
        private static readonly Dictionary<ElementOnPhaseStatus, string[]> NameHints = new Dictionary<ElementOnPhaseStatus, string[]>
        {
            { ElementOnPhaseStatus.Existing, new[] { "既存", "exist", "现有" } },
            { ElementOnPhaseStatus.Demolished, new[] { "解体", "demo", "拆除" } },
            { ElementOnPhaseStatus.New, new[] { "新規", "new", "新建" } },
            { ElementOnPhaseStatus.Temporary, new[] { "一時", "temp", "临时" } },
        };

        private const string NoneValue = "-";

        // 同じ Revit の中で、ライブ同期の書き出し（画面を出せない）でも同じ選択を使うため
        private static readonly Dictionary<string, Dictionary<ElementOnPhaseStatus, string>> Session =
            new Dictionary<string, Dictionary<ElementOnPhaseStatus, string>>();

        /// <summary>
        /// ビューで「上書き」表示になっている状態のマテリアルを決める。上書きが無ければ null。
        /// interactive のときは、まだ選んでいないモデル（または forceDialog）で選ぶ画面を出す。キャンセルなら cancelled = true。
        /// </summary>
        public static PhaseOverrides Resolve(Document doc, View view, ExternalCommandData commandData, bool interactive, bool forceDialog,
            out bool cancelled)
        {
            cancelled = false;
            List<ElementOnPhaseStatus> statuses = OverriddenStatuses(doc, view);
            if (statuses.Count == 0) return null;

            string key = ProjectKey(doc);
            Dictionary<ElementOnPhaseStatus, string> saved = LoadSaved(key);
            if (Session.TryGetValue(key, out var session))
                foreach (var kv in session) saved[kv.Key] = kv.Value;

            // 決まっていない状態は、名前で探した物を仮に入れる
            var choice = new Dictionary<ElementOnPhaseStatus, ElementId>();
            bool allSaved = true;
            foreach (ElementOnPhaseStatus s in statuses)
            {
                if (saved.TryGetValue(s, out string uid))
                {
                    choice[s] = uid == NoneValue ? ElementId.InvalidElementId : (doc.GetElement(uid) as Material)?.Id ?? FindByName(doc, s);
                    if (uid != NoneValue && doc.GetElement(uid) == null) allSaved = false;  // 覚えていたマテリアルが消えた
                }
                else
                {
                    choice[s] = FindByName(doc, s);
                    allSaved = false;
                }
            }

            if (interactive && (forceDialog || !allSaved))
            {
                var dialog = new PhaseMaterialDialog(doc, statuses, choice);
                if (commandData != null) dialog.SetRevitOwner(commandData);
                if (dialog.ShowDialog() != true)
                {
                    cancelled = true;
                    return null;
                }
                choice = dialog.Choice;
                var store = choice.ToDictionary(kv => kv.Key,
                    kv => kv.Value == ElementId.InvalidElementId ? NoneValue : doc.GetElement(kv.Value)?.UniqueId ?? NoneValue);
                Session[key] = store;
                Save(key, store);
            }

            var result = new PhaseOverrides { ViewPhaseId = view.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId() ?? ElementId.InvalidElementId };
            foreach (var kv in choice)
                if (kv.Value != ElementId.InvalidElementId) result.Materials[kv.Key] = kv.Value;
            DiagLog.Cmd("RealtimeRender", "フェーズの上書き: " +
                string.Join(", ", statuses.Select(s => $"{s}={(choice[s] == ElementId.InvalidElementId ? "なし" : (doc.GetElement(choice[s]) as Material)?.Name)}")));
            return result.ViewPhaseId == ElementId.InvalidElementId || result.Materials.Count == 0 ? null : result;
        }

        /// <summary>ビューのフェーズ フィルタで「上書き」表示になっている状態。</summary>
        private static List<ElementOnPhaseStatus> OverriddenStatuses(Document doc, View view)
        {
            var list = new List<ElementOnPhaseStatus>();
            ElementId filterId = view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER)?.AsElementId();
            if (!(doc.GetElement(filterId ?? ElementId.InvalidElementId) is PhaseFilter filter)) return list;
            foreach (ElementOnPhaseStatus s in Statuses)
            {
                try
                {
                    if (filter.GetPhaseStatusPresentation(s) == PhaseStatusPresentation.ShowOverriden) list.Add(s);
                }
                catch (Exception) { }
            }
            return list;
        }

        /// <summary>名前に「フェーズ」と状態（既存・解体など）の両方を含むマテリアルを探す。</summary>
        private static ElementId FindByName(Document doc, ElementOnPhaseStatus status)
        {
            string[] phaseWords = { "フェーズ", "phase", "阶段" };
            foreach (Material m in new FilteredElementCollector(doc).OfClass(typeof(Material)).Cast<Material>().OrderBy(m => m.Name))
            {
                string name = m.Name.ToLowerInvariant();
                if (phaseWords.Any(w => name.Contains(w)) && NameHints[status].Any(w => name.Contains(w))) return m.Id;
            }
            return ElementId.InvalidElementId;
        }

        private static string ProjectKey(Document doc)
        {
            string id = doc.ProjectInformation?.UniqueId;
            return string.IsNullOrEmpty(id) ? doc.Title : id;
        }

        private static string SettingsPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Tools28", "RealtimeRender", "phase-materials.txt");

        // 1 行 = モデル|状態|マテリアルの UniqueId（上書きしないなら "-"）
        private static Dictionary<ElementOnPhaseStatus, string> LoadSaved(string key)
        {
            var map = new Dictionary<ElementOnPhaseStatus, string>();
            try
            {
                if (!File.Exists(SettingsPath)) return map;
                foreach (string line in File.ReadAllLines(SettingsPath))
                {
                    string[] parts = line.Split('|');
                    if (parts.Length == 3 && parts[0] == key && Enum.TryParse(parts[1], out ElementOnPhaseStatus s)) map[s] = parts[2];
                }
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"フェーズのマテリアルの設定を読めませんでした: {ex.Message}");
            }
            return map;
        }

        private static void Save(string key, Dictionary<ElementOnPhaseStatus, string> map)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
                var lines = File.Exists(SettingsPath)
                    ? File.ReadAllLines(SettingsPath).Where(l => !l.StartsWith(key + "|")).ToList()
                    : new List<string>();
                lines.AddRange(map.Select(kv => $"{key}|{kv.Key}|{kv.Value}"));
                File.WriteAllLines(SettingsPath, lines);
            }
            catch (Exception ex)
            {
                DiagLog.Cmd("RealtimeRender", $"フェーズのマテリアルの設定を保存できませんでした: {ex.Message}");
            }
        }

        internal static string StatusLabel(ElementOnPhaseStatus s)
        {
            switch (s)
            {
                case ElementOnPhaseStatus.Existing: return Loc.S("RtRender.Phase.Existing");
                case ElementOnPhaseStatus.Demolished: return Loc.S("RtRender.Phase.Demolished");
                case ElementOnPhaseStatus.New: return Loc.S("RtRender.Phase.New");
                default: return Loc.S("RtRender.Phase.Temporary");
            }
        }
    }

    /// <summary>
    /// フェーズの状態ごとのマテリアルを選ぶ小さな画面（「上書き」表示の状態だけ並べる）。
    /// </summary>
    internal class PhaseMaterialDialog : Window
    {
        private readonly Dictionary<ElementOnPhaseStatus, ComboBox> _boxes = new Dictionary<ElementOnPhaseStatus, ComboBox>();
        private readonly List<KeyValuePair<string, ElementId>> _items;

        public Dictionary<ElementOnPhaseStatus, ElementId> Choice { get; } = new Dictionary<ElementOnPhaseStatus, ElementId>();

        public PhaseMaterialDialog(Document doc, List<ElementOnPhaseStatus> statuses, Dictionary<ElementOnPhaseStatus, ElementId> initial)
        {
            Title = Loc.S("RtRender.Phase.Title");
            SizeToContent = SizeToContent.WidthAndHeight;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            _items = new List<KeyValuePair<string, ElementId>> { new KeyValuePair<string, ElementId>(Loc.S("RtRender.Phase.None"), ElementId.InvalidElementId) };
            _items.AddRange(new FilteredElementCollector(doc).OfClass(typeof(Material)).Cast<Material>()
                .OrderBy(m => m.Name).Select(m => new KeyValuePair<string, ElementId>(m.Name, m.Id)));

            var root = new StackPanel { Margin = new Thickness(16) };
            root.Children.Add(new TextBlock
            {
                Text = Loc.S("RtRender.Phase.Message"),
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 420,
                Margin = new Thickness(0, 0, 0, 12),
            });
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(300) });
            foreach (ElementOnPhaseStatus s in statuses)
            {
                int row = grid.RowDefinitions.Count;
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var label = new TextBlock { Text = PhaseMaterials.StatusLabel(s), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 4, 12, 4) };
                Grid.SetRow(label, row);
                grid.Children.Add(label);
                var box = new ComboBox
                {
                    ItemsSource = _items,
                    DisplayMemberPath = "Key",
                    IsTextSearchEnabled = true,
                    Margin = new Thickness(0, 4, 0, 4),
                };
                ElementId selected = initial.TryGetValue(s, out ElementId id) ? id : ElementId.InvalidElementId;
                box.SelectedIndex = Math.Max(0, _items.FindIndex(i => i.Value == selected));
                Grid.SetRow(box, row);
                Grid.SetColumn(box, 1);
                grid.Children.Add(box);
                _boxes[s] = box;
            }
            root.Children.Add(grid);

            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
            var ok = new Button { Content = Loc.S("RtRender.Phase.Start"), IsDefault = true, MinWidth = 120, Padding = new Thickness(10, 4, 10, 4), Margin = new Thickness(0, 0, 8, 0) };
            ok.Click += (o, e) =>
            {
                foreach (var kv in _boxes)
                    Choice[kv.Key] = kv.Value.SelectedItem is KeyValuePair<string, ElementId> item ? item.Value : ElementId.InvalidElementId;
                DialogResult = true;
            };
            var cancel = new Button { Content = Loc.S("Common.Cancel"), IsCancel = true, MinWidth = 90, Padding = new Thickness(10, 4, 10, 4) };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            root.Children.Add(buttons);
            Content = root;
        }
    }
}
