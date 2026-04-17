using System;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;
using Tools28.Localization;

namespace Tools28
{
    public class Application : IExternalApplication
    {
        internal static PulldownButton LanguagePulldown { get; private set; }

        private static readonly List<RibbonPanel> _panels = new List<RibbonPanel>();
        private static readonly Dictionary<string, RibbonItem> _buttons = new Dictionary<string, RibbonItem>();

        private static readonly string[] _panelKeys = {
            "Ribbon.Panel.GridBubble", "Ribbon.Panel.SheetView", "Ribbon.Panel.3DView",
            "Ribbon.Panel.Annotation", "Ribbon.Panel.Structural", "Ribbon.Panel.Excel",
            "Ribbon.Panel.Settings"
        };

        private static readonly Dictionary<string, string> _buttonTextKeys = new Dictionary<string, string>
        {
            { "GridBubbleBoth", "Ribbon.GridBubble.Both" },
            { "GridBubbleLeft", "Ribbon.GridBubble.Left" },
            { "GridBubbleRight", "Ribbon.GridBubble.Right" },
            { "SheetCreation", "Ribbon.Sheet.Create" },
            { "ViewportPositionCopy", "Ribbon.Viewport.Copy" },
            { "ViewportPositionPaste", "Ribbon.Viewport.Paste" },
            { "CropBoxCopy", "Ribbon.CropBox.Copy" },
            { "CropBoxPaste", "Ribbon.CropBox.Paste" },
            { "ViewCopy", "Ribbon.View.Copy" },
            { "ViewPaste", "Ribbon.View.Paste" },
            { "SectionBoxCopy", "Ribbon.SectionBox.Copy" },
            { "SectionBoxPaste", "Ribbon.SectionBox.Paste" },
            { "RoomTagAutoCreator", "Ribbon.RoomTag" },
            { "FilledRegionSplitMerge2", "Ribbon.FilledRegion" },
            { "BeamUnderLevel", "Ribbon.BeamUnder" },
            { "BeamTopLevel", "Ribbon.BeamTop" },
            { "FireProtection", "Ribbon.FireProtection" },
            { "ExcelExport", "Ribbon.Excel.Export" },
            { "ExcelImport", "Ribbon.Excel.Import" },
            { "About", "Ribbon.Settings.About" },
            { "Manual", "Ribbon.Settings.Manual" },
        };

        private static readonly Dictionary<string, string> _buttonTipKeys = new Dictionary<string, string>
        {
            { "GridBubbleBoth", "Ribbon.GridBubble.Both.Tip" },
            { "GridBubbleLeft", "Ribbon.GridBubble.Left.Tip" },
            { "GridBubbleRight", "Ribbon.GridBubble.Right.Tip" },
            { "SheetCreation", "Ribbon.Sheet.Create.Tip" },
            { "ViewportPositionCopy", "Ribbon.Viewport.Copy.Tip" },
            { "ViewportPositionPaste", "Ribbon.Viewport.Paste.Tip" },
            { "CropBoxCopy", "Ribbon.CropBox.Copy.Tip" },
            { "CropBoxPaste", "Ribbon.CropBox.Paste.Tip" },
            { "ViewCopy", "Ribbon.View.Copy.Tip" },
            { "ViewPaste", "Ribbon.View.Paste.Tip" },
            { "SectionBoxCopy", "Ribbon.SectionBox.Copy.Tip" },
            { "SectionBoxPaste", "Ribbon.SectionBox.Paste.Tip" },
            { "RoomTagAutoCreator", "Ribbon.RoomTag.Tip" },
            { "FilledRegionSplitMerge2", "Ribbon.FilledRegion.Tip" },
            { "BeamUnderLevel", "Ribbon.BeamUnder.Tip" },
            { "BeamTopLevel", "Ribbon.BeamTop.Tip" },
            { "FireProtection", "Ribbon.FireProtection.Tip" },
            { "ExcelExport", "Ribbon.Excel.Export.Tip" },
            { "ExcelImport", "Ribbon.Excel.Import.Tip" },
            { "About", "Ribbon.Settings.About.Tip" },
            { "Manual", "Ribbon.Settings.Manual.Tip" },
        };

        public Result OnStartup(UIControlledApplication application)
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

            try
            {
                string debugLogPath = @"C:\temp\Tools28_debug.txt";
                try
                {
                    Directory.CreateDirectory(@"C:\temp");
                    var asm = Assembly.GetExecutingAssembly();
                    string logContent = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] OnStartup 開始\n"
                        + $"  DLL場所: {asm.Location}\n"
                        + $"  DLLバージョン: {asm.GetName().Version}\n"
                        + $"  DLL更新日時: {File.GetLastWriteTime(asm.Location):yyyy-MM-dd HH:mm:ss}\n";
                    File.AppendAllText(debugLogPath, logContent);
                }
                catch { }

                string tabName = "28 Tools";
                application.CreateRibbonTab(tabName);

                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                CreateGridBubblePanel(application, tabName, assemblyPath);
                CreateSheetViewPanel(application, tabName, assemblyPath);
                CreateThreeDViewPanel(application, tabName, assemblyPath);
                CreateAnnotationPanel(application, tabName, assemblyPath);
                CreateStructuralPanel(application, tabName, assemblyPath);
                CreateExcelPanel(application, tabName, assemblyPath);
                CreateSettingsPanel(application, tabName, assemblyPath);

                Loc.LanguageChanged += UpdateRibbonLanguage;
                UpdateRibbonLanguage();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show(Loc.S("Common.Error"), string.Format(Loc.S("App.StartupFailed"), ex.Message));
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            Loc.LanguageChanged -= UpdateRibbonLanguage;
            AppDomain.CurrentDomain.AssemblyResolve -= OnAssemblyResolve;
            return Result.Succeeded;
        }

        internal static void UpdateRibbonLanguage()
        {
            for (int i = 0; i < _panels.Count && i < _panelKeys.Length; i++)
            {
                try { _panels[i].Title = Loc.S(_panelKeys[i]); } catch { }
            }

            foreach (var kv in _buttonTextKeys)
            {
                if (_buttons.TryGetValue(kv.Key, out var btn))
                {
                    try { btn.ItemText = Loc.S(kv.Value); } catch { }
                }
            }

            foreach (var kv in _buttonTipKeys)
            {
                if (_buttons.TryGetValue(kv.Key, out var btn))
                {
                    try { btn.ToolTip = Loc.S(kv.Value); } catch { }
                }
            }

            if (LanguagePulldown != null)
            {
                try
                {
                    LanguagePulldown.ItemText = Loc.CurrentLang;
                    LanguagePulldown.ToolTip = Loc.S("Ribbon.Settings.Lang.Tip");
                } catch { }
            }
        }

        private Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                string assemblyName = new AssemblyName(args.Name).Name;
                string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string dllPath = Path.Combine(assemblyDir, assemblyName + ".dll");

                if (File.Exists(dllPath))
                    return Assembly.LoadFrom(dllPath);
            }
            catch { }
            return null;
        }

        private RibbonItem AddButton(RibbonPanel panel, PushButtonData data)
        {
            var item = panel.AddItem(data);
            _buttons[data.Name] = item;
            return item;
        }

        private void CreateGridBubblePanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, Loc.S("Ribbon.Panel.GridBubble"));
            _panels.Add(panel);

            var bothData = new PushButtonData("GridBubbleBoth", Loc.S("Ribbon.GridBubble.Both"), assemblyPath, "Tools28.Commands.GridBubble.ExecuteGridBubbleBothCommand");
            bothData.ToolTip = Loc.S("Ribbon.GridBubble.Both.Tip");
            bothData.LargeImage = LoadImage("both_96.png");
            AddButton(panel, bothData);

            var leftData = new PushButtonData("GridBubbleLeft", Loc.S("Ribbon.GridBubble.Left"), assemblyPath, "Tools28.Commands.GridBubble.ExecuteGridBubbleLeftCommand");
            leftData.ToolTip = Loc.S("Ribbon.GridBubble.Left.Tip");
            leftData.LargeImage = LoadImage("left_96.png");
            AddButton(panel, leftData);

            var rightData = new PushButtonData("GridBubbleRight", Loc.S("Ribbon.GridBubble.Right"), assemblyPath, "Tools28.Commands.GridBubble.ExecuteGridBubbleRightCommand");
            rightData.ToolTip = Loc.S("Ribbon.GridBubble.Right.Tip");
            rightData.LargeImage = LoadImage("right_96.png");
            AddButton(panel, rightData);
        }

        private void CreateSheetViewPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, Loc.S("Ribbon.Panel.SheetView"));
            _panels.Add(panel);

            var sheetData = new PushButtonData("SheetCreation", Loc.S("Ribbon.Sheet.Create"), assemblyPath, "Tools28.Commands.SheetCreation.ExecuteSheetCreationCommand");
            sheetData.ToolTip = Loc.S("Ribbon.Sheet.Create.Tip");
            sheetData.LargeImage = LoadImage("sheet_creation_96.png");
            AddButton(panel, sheetData);

            panel.AddSeparator();

            var vpCopyData = new PushButtonData("ViewportPositionCopy", Loc.S("Ribbon.Viewport.Copy"), assemblyPath, "Tools28.Commands.ViewportPosition.ExecuteViewportPositionCopyCommand");
            vpCopyData.ToolTip = Loc.S("Ribbon.Viewport.Copy.Tip");
            vpCopyData.LargeImage = LoadImage("viewport_copy_96.png");
            AddButton(panel, vpCopyData);

            var vpPasteData = new PushButtonData("ViewportPositionPaste", Loc.S("Ribbon.Viewport.Paste"), assemblyPath, "Tools28.Commands.ViewportPosition.ExecuteViewportPositionPasteCommand");
            vpPasteData.ToolTip = Loc.S("Ribbon.Viewport.Paste.Tip");
            vpPasteData.LargeImage = LoadImage("viewport_paste_96.png");
            AddButton(panel, vpPasteData);

            panel.AddSeparator();

            var cbCopyData = new PushButtonData("CropBoxCopy", Loc.S("Ribbon.CropBox.Copy"), assemblyPath, "Tools28.Commands.CropBoxCopy.ExecuteCropBoxCopyCommand");
            cbCopyData.ToolTip = Loc.S("Ribbon.CropBox.Copy.Tip");
            cbCopyData.LargeImage = LoadImage("cropbox_copy_96.png");
            AddButton(panel, cbCopyData);

            var cbPasteData = new PushButtonData("CropBoxPaste", Loc.S("Ribbon.CropBox.Paste"), assemblyPath, "Tools28.Commands.CropBoxCopy.ExecuteCropBoxPasteCommand");
            cbPasteData.ToolTip = Loc.S("Ribbon.CropBox.Paste.Tip");
            cbPasteData.LargeImage = LoadImage("cropbox_paste_96.png");
            AddButton(panel, cbPasteData);
        }

        private void CreateThreeDViewPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, Loc.S("Ribbon.Panel.3DView"));
            _panels.Add(panel);

            var viewCopyData = new PushButtonData("ViewCopy", Loc.S("Ribbon.View.Copy"), assemblyPath, "Tools28.Commands.ViewCopy.ExecuteViewCopyCommand");
            viewCopyData.ToolTip = Loc.S("Ribbon.View.Copy.Tip");
            viewCopyData.LargeImage = LoadImage("view_copy_96.png");
            AddButton(panel, viewCopyData);

            var viewPasteData = new PushButtonData("ViewPaste", Loc.S("Ribbon.View.Paste"), assemblyPath, "Tools28.Commands.ViewCopy.ExecuteViewPasteCommand");
            viewPasteData.ToolTip = Loc.S("Ribbon.View.Paste.Tip");
            viewPasteData.LargeImage = LoadImage("view_paste_96.png");
            AddButton(panel, viewPasteData);

            panel.AddSeparator();

            var sbCopyData = new PushButtonData("SectionBoxCopy", Loc.S("Ribbon.SectionBox.Copy"), assemblyPath, "Tools28.Commands.SectionBoxCopy.ExecuteSectionBoxCopyCommand");
            sbCopyData.ToolTip = Loc.S("Ribbon.SectionBox.Copy.Tip");
            sbCopyData.LargeImage = LoadImage("sectionbox_copy_96.png");
            AddButton(panel, sbCopyData);

            var sbPasteData = new PushButtonData("SectionBoxPaste", Loc.S("Ribbon.SectionBox.Paste"), assemblyPath, "Tools28.Commands.SectionBoxCopy.ExecuteSectionBoxPasteCommand");
            sbPasteData.ToolTip = Loc.S("Ribbon.SectionBox.Paste.Tip");
            sbPasteData.LargeImage = LoadImage("sectionbox_paste_96.png");
            AddButton(panel, sbPasteData);
        }

        private void CreateAnnotationPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, Loc.S("Ribbon.Panel.Annotation"));
            _panels.Add(panel);

            var roomTagData = new PushButtonData("RoomTagAutoCreator", Loc.S("Ribbon.RoomTag"), assemblyPath, "Tools28.Commands.RoomTagCreator.RoomTagAutoCreatorCommand");
            roomTagData.ToolTip = Loc.S("Ribbon.RoomTag.Tip");
            roomTagData.LargeImage = LoadImage("room_tag_96.png");
            AddButton(panel, roomTagData);

            panel.AddSeparator();

            var filledData = new PushButtonData("FilledRegionSplitMerge2", Loc.S("Ribbon.FilledRegion"), assemblyPath, "Tools28.Commands.FilledRegionSplitMerge.FilledRegionSplitMergeCommand");
            filledData.ToolTip = Loc.S("Ribbon.FilledRegion.Tip");
            filledData.LargeImage = LoadImage("filled_region_96.png");
            AddButton(panel, filledData);
        }

        private void CreateStructuralPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, Loc.S("Ribbon.Panel.Structural"));
            _panels.Add(panel);

            var beamUnderData = new PushButtonData("BeamUnderLevel", Loc.S("Ribbon.BeamUnder"), assemblyPath, "Tools28.Commands.BeamUnderLevel.BeamUnderLevelCommand");
            beamUnderData.ToolTip = Loc.S("Ribbon.BeamUnder.Tip");
            beamUnderData.LargeImage = LoadImage("beam_under_level_96.png");
            AddButton(panel, beamUnderData);

            var beamTopData = new PushButtonData("BeamTopLevel", Loc.S("Ribbon.BeamTop"), assemblyPath, "Tools28.Commands.BeamTopLevel.BeamTopLevelCommand");
            beamTopData.ToolTip = Loc.S("Ribbon.BeamTop.Tip");
            beamTopData.LargeImage = LoadImage("beam_top_level_96.png");
            AddButton(panel, beamTopData);

            panel.AddSeparator();

            var fireData = new PushButtonData("FireProtection", Loc.S("Ribbon.FireProtection"), assemblyPath, "Tools28.Commands.FireProtection.FireProtectionCommand");
            fireData.ToolTip = Loc.S("Ribbon.FireProtection.Tip");
            fireData.LargeImage = LoadImage("fire_protection_96.png");
            AddButton(panel, fireData);
        }

        private void CreateExcelPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, Loc.S("Ribbon.Panel.Excel"));
            _panels.Add(panel);

            var exportData = new PushButtonData("ExcelExport", Loc.S("Ribbon.Excel.Export"), assemblyPath, "Tools28.Commands.ExcelExportImport.ExcelExportCommand");
            exportData.ToolTip = Loc.S("Ribbon.Excel.Export.Tip");
            exportData.LargeImage = LoadImage("excel_export_96.png");
            AddButton(panel, exportData);

            var importData = new PushButtonData("ExcelImport", Loc.S("Ribbon.Excel.Import"), assemblyPath, "Tools28.Commands.ExcelExportImport.ExcelImportCommand");
            importData.ToolTip = Loc.S("Ribbon.Excel.Import.Tip");
            importData.LargeImage = LoadImage("excel_import_96.png");
            AddButton(panel, importData);
        }

        private void CreateSettingsPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, Loc.S("Ribbon.Panel.Settings"));
            _panels.Add(panel);

            string currentLang = Loc.CurrentLang;
            string flagIcon = $"flag_{currentLang.ToLower()}_16.png";
            if (currentLang == "US") flagIcon = "flag_us_16.png";

            var langData = new PulldownButtonData("LanguageSwitch", currentLang);
            langData.ToolTip = Loc.S("Ribbon.Settings.Lang.Tip");
            langData.Image = LoadImage(flagIcon);

            var aboutData = new PushButtonData("About", Loc.S("Ribbon.Settings.About"), assemblyPath, "Tools28.Commands.LanguageSwitch.AboutCommand");
            aboutData.ToolTip = Loc.S("Ribbon.Settings.About.Tip");
            aboutData.Image = LoadImage("ver_16.png");

            var manualData = new PushButtonData("Manual", Loc.S("Ribbon.Settings.Manual"), assemblyPath, "Tools28.Commands.LanguageSwitch.ManualCommand");
            manualData.ToolTip = Loc.S("Ribbon.Settings.Manual.Tip");
            manualData.Image = LoadImage("manual_16.png");

            var stackedItems = panel.AddStackedItems(langData, aboutData, manualData);

            _buttons["About"] = stackedItems[1] as RibbonItem;
            _buttons["Manual"] = stackedItems[2] as RibbonItem;

            LanguagePulldown = stackedItems[0] as PulldownButton;
            if (LanguagePulldown != null)
            {
                var jpData = new PushButtonData("LangJP", "JP", assemblyPath, "Tools28.Commands.LanguageSwitch.SwitchToJapaneseCommand");
                jpData.Image = LoadImage("flag_jp_16.png");
                jpData.LargeImage = LoadImage("flag_jp_32.png");
                jpData.ToolTip = Loc.S("Ribbon.Settings.Lang.JP.Tip");
                LanguagePulldown.AddPushButton(jpData);

                var enData = new PushButtonData("LangUS", "US", assemblyPath, "Tools28.Commands.LanguageSwitch.SwitchToEnglishCommand");
                enData.Image = LoadImage("flag_us_16.png");
                enData.LargeImage = LoadImage("flag_us_32.png");
                enData.ToolTip = Loc.S("Ribbon.Settings.Lang.US.Tip");
                LanguagePulldown.AddPushButton(enData);

                var cnData = new PushButtonData("LangCN", "CN", assemblyPath, "Tools28.Commands.LanguageSwitch.SwitchToChineseCommand");
                cnData.Image = LoadImage("flag_cn_16.png");
                cnData.LargeImage = LoadImage("flag_cn_32.png");
                cnData.ToolTip = Loc.S("Ribbon.Settings.Lang.CN.Tip");
                LanguagePulldown.AddPushButton(cnData);
            }
        }

        private BitmapImage LoadImage(string fileName)
        {
            try
            {
                string packUri = $"pack://application:,,,/Tools28;component/Resources/Icons/{fileName}";
                var image = new BitmapImage(new Uri(packUri));
                image.Freeze();
                return image;
            }
            catch { }

            try
            {
                string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string filePath = Path.Combine(assemblyDir, "Icons", fileName);
                if (File.Exists(filePath))
                {
                    var image = new BitmapImage();
                    image.BeginInit();
                    image.UriSource = new Uri(filePath, UriKind.Absolute);
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    image.Freeze();
                    return image;
                }
            }
            catch { }

            return null;
        }
    }
}
