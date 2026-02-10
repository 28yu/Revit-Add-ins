using System;
using Autodesk.Revit.UI;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

namespace Tools28
{
    public class Application : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "28 Tools";
                application.CreateRibbonTab(tabName);

                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                // パネル作成
                CreateGridBubblePanel(application, tabName, assemblyPath);
                CreateSheetPanel(application, tabName, assemblyPath);
                CreateThreeDViewPanel(application, tabName, assemblyPath);
                CreateViewPanel(application, tabName, assemblyPath);
                CreateDetailPanel(application, tabName, assemblyPath);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("エラー", $"アドインの起動に失敗しました。\n{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// 符号ON/OFFパネルを作成
        /// </summary>
        private void CreateGridBubblePanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "符号 ON/OFF");

            // 両端ボタン
            PushButtonData bothButtonData = new PushButtonData(
                "GridBubbleBoth",
                "両端",
                assemblyPath,
                "Tools28.Commands.GridBubble.ExecuteGridBubbleBothCommand");
            bothButtonData.ToolTip = "通り芯・レベルの符号を両端に表示します";
            bothButtonData.LargeImage = LoadImage("both_32.png");
            panel.AddItem(bothButtonData);

            // 左端ボタン
            PushButtonData leftButtonData = new PushButtonData(
                "GridBubbleLeft",
                "左端",
                assemblyPath,
                "Tools28.Commands.GridBubble.ExecuteGridBubbleLeftCommand");
            leftButtonData.ToolTip = "通り芯・レベルの符号を左端のみに表示します";
            leftButtonData.LargeImage = LoadImage("left_32.png");
            panel.AddItem(leftButtonData);

            // 右端ボタン
            PushButtonData rightButtonData = new PushButtonData(
                "GridBubbleRight",
                "右端",
                assemblyPath,
                "Tools28.Commands.GridBubble.ExecuteGridBubbleRightCommand");
            rightButtonData.ToolTip = "通り芯・レベルの符号を右端のみに表示します";
            rightButtonData.LargeImage = LoadImage("right_32.png");
            panel.AddItem(rightButtonData);
        }

        /// <summary>
        /// シートパネルを作成
        /// </summary>
        private void CreateSheetPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "シート");

            // シート一括作成ボタン
            PushButtonData sheetButtonData = new PushButtonData(
                "SheetCreation",
                "一括作成",
                assemblyPath,
                "Tools28.Commands.SheetCreation.ExecuteSheetCreationCommand");
            sheetButtonData.ToolTip = "図枠を指定して複数のシートを一括作成します";
            sheetButtonData.LargeImage = LoadImage("sheet_creation_32.png");
            panel.AddItem(sheetButtonData);

            // セパレータ
            panel.AddSeparator();

            // ビューポート位置コピーボタン
            PushButtonData viewportCopyButtonData = new PushButtonData(
                "ViewportPositionCopy",
                "ビューポート\n位置コピー",
                assemblyPath,
                "Tools28.Commands.ViewportPosition.ExecuteViewportPositionCopyCommand");
            viewportCopyButtonData.ToolTip = "シート上のビューポート位置をコピーします";
            viewportCopyButtonData.LargeImage = LoadImage("viewport_copy_32.png");
            panel.AddItem(viewportCopyButtonData);

            // ビューポート位置ペーストボタン
            PushButtonData viewportPasteButtonData = new PushButtonData(
                "ViewportPositionPaste",
                "ビューポート\n位置ペースト",
                assemblyPath,
                "Tools28.Commands.ViewportPosition.ExecuteViewportPositionPasteCommand");
            viewportPasteButtonData.ToolTip = "コピーしたビューポート位置を他のシートに適用します";
            viewportPasteButtonData.LargeImage = LoadImage("viewport_paste_32.png");
            panel.AddItem(viewportPasteButtonData);
        }

        /// <summary>
        /// 3Dビューパネルを作成
        /// </summary>
        private void CreateThreeDViewPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "3Dビュー");

            // 視点コピーボタン
            PushButtonData viewCopyButtonData = new PushButtonData(
                "ViewCopy",
                "視点コピー",
                assemblyPath,
                "Tools28.Commands.ViewCopy.ExecuteViewCopyCommand");
            viewCopyButtonData.ToolTip = "3Dビューの視点をコピーします";
            viewCopyButtonData.LargeImage = LoadImage("view_copy_32.png");
            panel.AddItem(viewCopyButtonData);

            // 視点ペーストボタン
            PushButtonData viewPasteButtonData = new PushButtonData(
                "ViewPaste",
                "視点ペースト",
                assemblyPath,
                "Tools28.Commands.ViewCopy.ExecuteViewPasteCommand");
            viewPasteButtonData.ToolTip = "コピーした視点を他の3Dビューに適用します";
            viewPasteButtonData.LargeImage = LoadImage("view_paste_32.png");
            panel.AddItem(viewPasteButtonData);

            // セパレータ
            panel.AddSeparator();

            // 切断ボックスコピーボタン
            PushButtonData sectionBoxCopyButtonData = new PushButtonData(
                "SectionBoxCopy",
                "切断ボックス\nコピー",
                assemblyPath,
                "Tools28.Commands.SectionBoxCopy.ExecuteSectionBoxCopyCommand");
            sectionBoxCopyButtonData.ToolTip = "3Dビューの切断ボックスをコピーします";
            sectionBoxCopyButtonData.LargeImage = LoadImage("sectionbox_copy_32.png");
            panel.AddItem(sectionBoxCopyButtonData);

            // 切断ボックスペーストボタン
            PushButtonData sectionBoxPasteButtonData = new PushButtonData(
                "SectionBoxPaste",
                "切断ボックス\nペースト",
                assemblyPath,
                "Tools28.Commands.SectionBoxCopy.ExecuteSectionBoxPasteCommand");
            sectionBoxPasteButtonData.ToolTip = "コピーした切断ボックスを他の3Dビューに適用します";
            sectionBoxPasteButtonData.LargeImage = LoadImage("sectionbox_paste_32.png");
            panel.AddItem(sectionBoxPasteButtonData);
        }

        /// <summary>
        /// ビューパネルを作成
        /// </summary>
        private void CreateViewPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "ビュー");

            // トリミング領域コピーボタン
            PushButtonData cropBoxCopyButtonData = new PushButtonData(
                "CropBoxCopy",
                "トリミング\n領域コピー",
                assemblyPath,
                "Tools28.Commands.CropBoxCopy.ExecuteCropBoxCopyCommand");
            cropBoxCopyButtonData.ToolTip = "ビューのトリミング領域をコピーします";
            cropBoxCopyButtonData.LargeImage = LoadImage("cropbox_copy_32.png");
            panel.AddItem(cropBoxCopyButtonData);

            // トリミング領域ペーストボタン
            PushButtonData cropBoxPasteButtonData = new PushButtonData(
                "CropBoxPaste",
                "トリミング\n領域ペースト",
                assemblyPath,
                "Tools28.Commands.CropBoxCopy.ExecuteCropBoxPasteCommand");
            cropBoxPasteButtonData.ToolTip = "コピーしたトリミング領域を他のビューに適用します";
            cropBoxPasteButtonData.LargeImage = LoadImage("cropbox_paste_32.png");
            panel.AddItem(cropBoxPasteButtonData);
        }

        /// <summary>
        /// 詳細パネルを作成
        /// </summary>
        private void CreateDetailPanel(UIControlledApplication application, string tabName, string assemblyPath)
        {
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "詳細");

            // 塗潰し領域 分割/統合ボタン（一時的に無効化：ビルドエラー回避のため）
            /*
            PushButtonData filledRegionButtonData = new PushButtonData(
                "FilledRegionSplitMerge",
                "領域",
                assemblyPath,
                "Tools28.Commands.FilledRegionSplitMerge.FilledRegionSplitMergeCommand");
            filledRegionButtonData.ToolTip = "配置されている塗潰領域を分割/統合します";
            filledRegionButtonData.LargeImage = LoadImage("filled_region_32.png");
            panel.AddItem(filledRegionButtonData);
            */
        }

        /// <summary>
        /// 画像を読み込み（ハイブリッド方式）
        /// </summary>
        private BitmapImage LoadImage(string fileName)
        {
            BitmapImage image = null;

            // 方法1: リソースから読み込み（優先）
            try
            {
                string packUri = $"pack://application:,,,/Tools28;component/Resources/Icons/{fileName}";
                image = new BitmapImage(new Uri(packUri));
                image.Freeze();
                return image;
            }
            catch
            {
                // リソースからの読み込みに失敗
            }

            // 方法2: 外部ファイルから読み込み（フォールバック）
            try
            {
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyPath);
                string iconFolder = Path.Combine(assemblyDir, "Icons");
                string filePath = Path.Combine(iconFolder, fileName);

                if (File.Exists(filePath))
                {
                    image = new BitmapImage();
                    image.BeginInit();
                    image.UriSource = new Uri(filePath, UriKind.Absolute);
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    image.Freeze();
                    return image;
                }
            }
            catch
            {
                // 外部ファイルからの読み込みにも失敗
            }

            return null;
        }
    }
}