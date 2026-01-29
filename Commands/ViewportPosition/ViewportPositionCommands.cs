using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace Tools28.Commands.ViewportPosition
{
    // ビューポート位置情報を保持する静的クラス
    public static class ViewportPositionClipboard
    {
        public static XYZ CopiedPosition { get; set; }
        public static string SourceViewportInfo { get; set; }
        public static string SourceSheetName { get; set; }
        public static bool HasCopiedPosition => CopiedPosition != null;

        public static void Clear()
        {
            CopiedPosition = null;
            SourceViewportInfo = null;
            SourceSheetName = null;
        }
    }

    // ビューマッチング用のヘルパークラス
    public static class ViewMatcher
    {
        public enum MatchingMode
        {
            ExactMatch,      // 完全一致
            PatternMatch,    // パターンマッチング（改善版）
            FloorMatch       // 階層マッチング（階層を保持）
        }

        public static List<(Viewport source, Viewport target, string matchReason)> FindMatches(
            List<Viewport> sourceViewports, List<Viewport> targetViewports, Document doc, MatchingMode mode)
        {
            var matches = new List<(Viewport source, Viewport target, string matchReason)>();

            foreach (var sourceVp in sourceViewports)
            {
                var sourceView = doc.GetElement(sourceVp.ViewId) as View;
                if (sourceView == null) continue;

                var targetVp = FindBestMatch(sourceView, targetViewports, doc, mode);
                if (targetVp.viewport != null)
                {
                    matches.Add((sourceVp, targetVp.viewport, targetVp.reason));
                }
            }

            return matches;
        }

        private static (Viewport viewport, string reason) FindBestMatch(
            View sourceView, List<Viewport> targetViewports, Document doc, MatchingMode mode)
        {
            string sourceName = sourceView.Name;

            foreach (var targetVp in targetViewports)
            {
                var targetView = doc.GetElement(targetVp.ViewId) as View;
                if (targetView == null) continue;

                string targetName = targetView.Name;

                switch (mode)
                {
                    case MatchingMode.ExactMatch:
                        if (sourceName.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                        {
                            return (targetVp, "完全一致");
                        }
                        break;

                    case MatchingMode.PatternMatch:
                        var patternMatch = CheckEnhancedPatternMatch(sourceName, targetName);
                        if (patternMatch.isMatch)
                        {
                            return (targetVp, patternMatch.reason);
                        }
                        break;

                    case MatchingMode.FloorMatch:
                        var floorMatch = CheckFloorLevelMatch(sourceName, targetName);
                        if (floorMatch.isMatch)
                        {
                            return (targetVp, floorMatch.reason);
                        }
                        break;
                }
            }

            return (null, null);
        }

        private static (bool isMatch, string reason) CheckEnhancedPatternMatch(string sourceName, string targetName)
        {
            // 1. まず番号パターンをチェック（test1, test2など）- 新規追加
            var numberMatch = CheckNumberPatternMatch(sourceName, targetName);
            if (numberMatch.isMatch)
            {
                return numberMatch;
            }

            // 2. ビュータイプパターン（既存のコード）
            var viewTypePatterns = new Dictionary<string, string[]>
            {
                ["平面図系"] = new[] { "平面図", "平面", "PLAN", "フロアプラン", "FLOOR PLAN" },
                ["断面図系"] = new[] { "断面図", "断面", "SECTION", "セクション", "矩計図", "矩計", "WALL SECTION" },
                ["立面図系"] = new[] { "立面図", "立面", "ELEVATION", "エレベーション", "外観図" },
                ["詳細図系"] = new[] { "詳細図", "詳細", "DETAIL", "ディテール", "部分詳細", "拡大図" },
                ["天井図系"] = new[] { "天井伏図", "天井", "CEILING", "天井プラン", "CEILING PLAN" },
                ["配置図系"] = new[] { "配置図", "配置", "SITE", "サイトプラン", "SITE PLAN", "外構図" },
                ["展開図系"] = new[] { "展開図", "展開", "INTERIOR", "内観図", "室内展開" },
                ["構造図系"] = new[] { "構造図", "構造", "STRUCTURAL", "梁伏図", "基礎図", "躯体図" },
                ["設備図系"] = new[] { "設備図", "設備", "MEP", "機械図", "電気図", "空調図", "衛生図" }
            };

            // 方位・位置パターン
            var orientationPatterns = new[] { "東", "西", "南", "北", "E", "W", "S", "N", "EAST", "WEST", "SOUTH", "NORTH" };

            // 番号パターン（01, 02, A, B, ①, ②など）
            var numberPatterns = new[] { @"\d+", @"[A-Z]", @"[①-⑩]", @"その\d+", @"TYPE\s*[A-Z]", @"タイプ\s*[A-Z]" };

            // ビュータイプでマッチング
            foreach (var viewType in viewTypePatterns)
            {
                bool sourceHasType = viewType.Value.Any(pattern =>
                    sourceName.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0);
                bool targetHasType = viewType.Value.Any(pattern =>
                    targetName.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0);

                if (sourceHasType && targetHasType)
                {
                    // 追加の共通要素をチェック
                    var commonElements = new List<string>();

                    // 方位チェック
                    foreach (var orientation in orientationPatterns)
                    {
                        if (sourceName.IndexOf(orientation, StringComparison.OrdinalIgnoreCase) >= 0 &&
                            targetName.IndexOf(orientation, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            commonElements.Add($"方位:{orientation}");
                        }
                    }

                    // 番号パターンチェック
                    foreach (var numPattern in numberPatterns)
                    {
                        var sourceMatch = Regex.Match(sourceName, numPattern);
                        var targetMatch = Regex.Match(targetName, numPattern);

                        if (sourceMatch.Success && targetMatch.Success &&
                            sourceMatch.Value == targetMatch.Value)
                        {
                            commonElements.Add($"番号:{sourceMatch.Value}");
                        }
                    }

                    // 共通キーワードチェック（カスタムワード）
                    var commonKeywords = ExtractCommonKeywords(sourceName, targetName);
                    if (commonKeywords.Count > 0)
                    {
                        commonElements.AddRange(commonKeywords.Select(k => $"共通:{k}"));
                    }

                    // マッチング理由の構築
                    string reason = $"{viewType.Key}";
                    if (commonElements.Count > 0)
                    {
                        reason += $" + {string.Join(", ", commonElements)}";
                    }

                    return (true, reason);
                }
            }

            // 3. ビュータイプが一致しない場合でも、複数の共通要素があればマッチ
            var sharedElements = ExtractSharedElements(sourceName, targetName);
            if (sharedElements.Count >= 2) // 2つ以上の共通要素
            {
                return (true, $"共通要素: {string.Join(", ", sharedElements)}");
            }

            return (false, null);
        }

        // 新規追加: 番号パターンマッチング
        private static (bool isMatch, string reason) CheckNumberPatternMatch(string sourceName, string targetName)
        {
            // 様々な番号パターンに対応
            var patterns = new[]
            {
                // パターン1: 末尾の数字（test1, test2）
                new { Pattern = @"^(.+?)(\d+)$", Description = "末尾番号" },
                
                // パターン2: 区切り文字付き末尾の数字（test-1, test_2, test 3）
                new { Pattern = @"^(.+?)[-_\s]+(\d+)$", Description = "区切り付き末尾番号" },
                
                // パターン3: 括弧内の数字（test(1), test(2)）
                new { Pattern = @"^(.+?)\((\d+)\)(.*)$", Description = "括弧内番号" },
                
                // パターン4: 中間の数字（test1view, test2view）
                new { Pattern = @"^(.+?)(\d+)(.+)$", Description = "中間番号" },
                
                // パターン5: アルファベット番号（testA, testB）
                new { Pattern = @"^(.+?)([A-Z])$", Description = "アルファベット番号" },
                
                // パターン6: ローマ数字（testⅠ, testⅡ）
                new { Pattern = @"^(.+?)([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹ]+)$", Description = "ローマ数字" },
                
                // パターン7: ゼロパディング（test01, test02）
                new { Pattern = @"^(.+?)(\d{2,})$", Description = "ゼロパディング番号" },
                
                // パターン8: 日本語番号（その1, その2）
                new { Pattern = @"^(.+?)(その|第|No\.?|№)(\d+)(.*)$", Description = "日本語番号" }
            };

            foreach (var patternInfo in patterns)
            {
                var sourceMatch = Regex.Match(sourceName, patternInfo.Pattern, RegexOptions.IgnoreCase);
                var targetMatch = Regex.Match(targetName, patternInfo.Pattern, RegexOptions.IgnoreCase);

                if (sourceMatch.Success && targetMatch.Success)
                {
                    // グループ数に応じて処理を分岐
                    string sourceBase, targetBase, sourceNum, targetNum;

                    if (sourceMatch.Groups.Count == 3) // 単純なパターン（前後2つ）
                    {
                        sourceBase = sourceMatch.Groups[1].Value;
                        targetBase = targetMatch.Groups[1].Value;
                        sourceNum = sourceMatch.Groups[2].Value;
                        targetNum = targetMatch.Groups[2].Value;
                    }
                    else if (sourceMatch.Groups.Count == 4) // 中間番号パターン（前中後3つ）
                    {
                        sourceBase = sourceMatch.Groups[1].Value + sourceMatch.Groups[3].Value;
                        targetBase = targetMatch.Groups[1].Value + targetMatch.Groups[3].Value;
                        sourceNum = sourceMatch.Groups[2].Value;
                        targetNum = targetMatch.Groups[2].Value;
                    }
                    else if (sourceMatch.Groups.Count == 5) // 日本語番号パターン（4つ）
                    {
                        sourceBase = sourceMatch.Groups[1].Value + sourceMatch.Groups[2].Value + sourceMatch.Groups[4].Value;
                        targetBase = targetMatch.Groups[1].Value + targetMatch.Groups[2].Value + targetMatch.Groups[4].Value;
                        sourceNum = sourceMatch.Groups[3].Value;
                        targetNum = targetMatch.Groups[3].Value;
                    }
                    else
                    {
                        continue;
                    }

                    // ベース部分が同じかチェック（大文字小文字を無視）
                    if (sourceBase.Equals(targetBase, StringComparison.OrdinalIgnoreCase))
                    {
                        // 番号が異なることを確認（同じ番号はスキップ）
                        if (!sourceNum.Equals(targetNum))
                        {
                            return (true, $"{patternInfo.Description}: {sourceBase}[{sourceNum}→{targetNum}]");
                        }
                    }
                }
            }

            // 特殊ケース: 番号の前後が入れ替わっているパターン（1-test と test-2）
            var reversePattern1 = Regex.Match(sourceName, @"^(\d+)[-_\s]*(.+)$");
            var reversePattern2 = Regex.Match(targetName, @"^(.+?)[-_\s]*(\d+)$");

            if (reversePattern1.Success && reversePattern2.Success)
            {
                if (reversePattern1.Groups[2].Value.Equals(reversePattern2.Groups[1].Value, StringComparison.OrdinalIgnoreCase))
                {
                    return (true, $"番号位置違い: {reversePattern1.Groups[2].Value}");
                }
            }

            return (false, null);
        }

        private static List<string> ExtractCommonKeywords(string source, string target)
        {
            var commonKeywords = new List<string>();

            // よく使われるキーワード
            var keywords = new[]
            {
                "既存", "新設", "改修", "撤去",
                "before", "after", "BEFORE", "AFTER",
                "現況", "計画", "将来",
                "共用", "専用", "住戸", "店舗",
                "エントランス", "ロビー", "廊下", "階段",
                "A棟", "B棟", "C棟", "本館", "別館",
                "ZONE", "ゾーン", "エリア", "AREA"
            };

            foreach (var keyword in keywords)
            {
                if (source.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 &&
                    target.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    commonKeywords.Add(keyword);
                }
            }

            return commonKeywords;
        }

        private static List<string> ExtractSharedElements(string source, string target)
        {
            var sharedElements = new List<string>();

            // 単語に分割（スペース、ハイフン、アンダースコアで分割）
            var separators = new[] { ' ', '-', '_', '・', '／', '(' };
            var sourceWords = source.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            var targetWords = target.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            // 3文字以上の共通単語を抽出
            foreach (var sourceWord in sourceWords)
            {
                if (sourceWord.Length >= 3)
                {
                    foreach (var targetWord in targetWords)
                    {
                        if (sourceWord.Equals(targetWord, StringComparison.OrdinalIgnoreCase))
                        {
                            sharedElements.Add(sourceWord);
                            break;
                        }
                    }
                }
            }

            return sharedElements.Distinct().ToList();
        }

        private static (bool isMatch, string reason) CheckFloorLevelMatch(string sourceName, string targetName)
        {
            // 階層パターンを抽出して比較
            var floorInfo = ExtractFloorInfo(sourceName);
            var targetFloorInfo = ExtractFloorInfo(targetName);

            if (floorInfo.hasFloor && targetFloorInfo.hasFloor)
            {
                // 同じ階層タイプ（地上階、地下階、屋上など）
                if (floorInfo.floorType == targetFloorInfo.floorType)
                {
                    // ベース名（階層以外の部分）も確認
                    if (floorInfo.baseName.Equals(targetFloorInfo.baseName, StringComparison.OrdinalIgnoreCase))
                    {
                        return (true, $"階層タイプ一致: {floorInfo.floorType} ({floorInfo.floor} → {targetFloorInfo.floor})");
                    }

                    // ベース名に共通要素があるかチェック
                    var commonBase = ExtractSharedElements(floorInfo.baseName, targetFloorInfo.baseName);
                    if (commonBase.Count > 0)
                    {
                        return (true, $"階層タイプ一致: {floorInfo.floorType} + 共通:{string.Join(",", commonBase)}");
                    }
                }

                // 階層の数値的な関係性をチェック（連続階など）
                if (floorInfo.floorNumber.HasValue && targetFloorInfo.floorNumber.HasValue)
                {
                    int diff = Math.Abs(floorInfo.floorNumber.Value - targetFloorInfo.floorNumber.Value);
                    if (diff == 1 && floorInfo.baseName.Equals(targetFloorInfo.baseName, StringComparison.OrdinalIgnoreCase))
                    {
                        return (true, $"連続階: {floorInfo.floor} ⇔ {targetFloorInfo.floor}");
                    }
                }
            }

            return (false, null);
        }

        private static (bool hasFloor, string floor, string floorType, int? floorNumber, string baseName) ExtractFloorInfo(string name)
        {
            // 階層パターンの定義
            var patterns = new[]
            {
                (@"(\d+)F", "地上階"),
                (@"(\d+)階", "地上階"),
                (@"B(\d+)", "地下階"),
                (@"地下(\d+)階", "地下階"),
                (@"RF", "屋上階"),
                (@"屋上", "屋上階"),
                (@"PH", "ペントハウス"),
                (@"(\d+)FL", "地上階"),
                (@"GL", "地上階"),
                (@"(\d+)st", "地上階"),
                (@"(\d+)nd", "地上階"),
                (@"(\d+)rd", "地上階"),
                (@"(\d+)th", "地上階")
            };

            foreach (var (pattern, type) in patterns)
            {
                var match = Regex.Match(name, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string floor = match.Value;
                    string baseName = name.Replace(match.Value, "").Trim().Trim('-', '_', ' ');

                    // 階数を抽出
                    int? floorNumber = null;
                    var numberMatch = Regex.Match(match.Value, @"\d+");
                    if (numberMatch.Success)
                    {
                        floorNumber = int.Parse(numberMatch.Value);
                    }

                    return (true, floor, type, floorNumber, baseName);
                }
            }

            return (false, null, null, null, name);
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ExecuteViewportPositionCopyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // アクティブビューがシートかチェック
                ViewSheet activeSheet = doc.ActiveView as ViewSheet;
                if (activeSheet == null)
                {
                    message = "シートをアクティブにしてから実行してください。";
                    return Result.Failed;
                }

                // 選択されたビューポートを取得
                var selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 0)
                {
                    message = "ビューポートを選択してから実行してください。";
                    return Result.Failed;
                }

                // 選択された要素からビューポートを抽出
                var viewports = selectedIds
                    .Select(id => doc.GetElement(id))
                    .OfType<Viewport>()
                    .ToList();

                if (viewports.Count == 0)
                {
                    message = "選択された要素にビューポートが含まれていません。シート上のビューポートを選択してください。";
                    return Result.Failed;
                }

                if (viewports.Count > 1)
                {
                    message = "複数のビューポートが選択されています。1つのビューポートを選択してください。";
                    return Result.Failed;
                }

                Viewport selectedViewport = viewports[0];

                // 位置情報を取得
                XYZ position = selectedViewport.GetBoxCenter();

                // 関連するビュー情報を取得
                View relatedView = doc.GetElement(selectedViewport.ViewId) as View;
                string viewInfo = relatedView != null ? relatedView.Name : "不明なビュー";

                // クリップボードに保存
                ViewportPositionClipboard.CopiedPosition = position;
                ViewportPositionClipboard.SourceViewportInfo = viewInfo;
                ViewportPositionClipboard.SourceSheetName = activeSheet.Name;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"位置コピー中にエラーが発生しました。{ex.Message}";
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ExecuteViewportPositionPasteCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // コピーされた位置情報があるかチェック
                if (!ViewportPositionClipboard.HasCopiedPosition)
                {
                    message = "コピーされた位置情報がありません。先に位置コピーを実行してください。";
                    return Result.Failed;
                }

                // 選択方法を判定
                var selectedIds = uidoc.Selection.GetElementIds();

                // パターン1: アクティブシート上でビューポートが選択されている場合
                ViewSheet activeSheet = doc.ActiveView as ViewSheet;
                if (activeSheet != null && selectedIds.Count > 0)
                {
                    var viewportsOnSheet = selectedIds
                        .Select(id => doc.GetElement(id))
                        .OfType<Viewport>()
                        .ToList();

                    if (viewportsOnSheet.Count > 0)
                    {
                        return ExecuteDirectViewportPasteImpl(doc, activeSheet, viewportsOnSheet);
                    }
                }

                // パターン2: プロジェクトブラウザでの選択を分析
                if (selectedIds.Count > 0)
                {
                    var selectedElements = selectedIds.Select(id => doc.GetElement(id)).ToList();

                    // シートが選択されている場合
                    var selectedSheets = selectedElements.OfType<ViewSheet>().ToList();
                    if (selectedSheets.Count > 0)
                    {
                        // 複数シート選択の場合は自動マッチング機能を使用
                        if (selectedSheets.Count > 1)
                        {
                            return ExecuteAutoMatchingImpl(doc, selectedSheets, ViewMatcher.MatchingMode.ExactMatch, "完全一致");
                        }
                        else
                        {
                            return ExecuteMultiSheetPasteImpl(doc, selectedSheets);
                        }
                    }

                    // ビューが選択されている場合
                    var selectedViews = selectedElements.OfType<View>()
                        .Where(v => !(v is ViewSheet))
                        .ToList();

                    if (selectedViews.Count > 0)
                    {
                        return ExecuteViewBasedPasteImpl(doc, selectedViews);
                    }
                }

                // どれでもない場合はエラー
                message = "ビューポート、シート、またはビューを選択してから実行してください。";
                return Result.Failed;
            }
            catch (Exception ex)
            {
                message = $"位置ペースト中にエラーが発生しました。{ex.Message}";
                return Result.Failed;
            }
        }

        private Result ExecuteAutoMatchingImpl(Document doc, List<ViewSheet> targetSheets,
            ViewMatcher.MatchingMode mode, string modeDescription)
        {
            // コピー元シートの情報を取得
            ViewSheet sourceSheet = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .FirstOrDefault(s => s.Name == ViewportPositionClipboard.SourceSheetName);

            if (sourceSheet == null)
            {
                return Result.Failed;
            }

            // コピー元シートのビューポートを取得
            var sourceViewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .WhereElementIsNotElementType()
                .Cast<Viewport>()
                .Where(vp => vp.SheetId == sourceSheet.Id)
                .ToList();

            if (sourceViewports.Count == 0)
            {
                return Result.Failed;
            }

            // 各対象シートでマッチングを実行
            var allMatches = new List<(ViewSheet sheet, List<(Viewport source, Viewport target, string reason)> matches)>();
            int totalMatches = 0;

            foreach (var targetSheet in targetSheets)
            {
                if (targetSheet.Id == sourceSheet.Id) continue; // 同じシートはスキップ

                var targetViewports = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .WhereElementIsNotElementType()
                    .Cast<Viewport>()
                    .Where(vp => vp.SheetId == targetSheet.Id)
                    .ToList();

                var matches = ViewMatcher.FindMatches(sourceViewports, targetViewports, doc, mode);
                if (matches.Count > 0)
                {
                    allMatches.Add((targetSheet, matches));
                    totalMatches += matches.Count;
                }
            }

            if (totalMatches == 0)
            {
                return Result.Failed;
            }

            // 実際の位置変更を実行
            return ApplyAutoMatchingResultsImpl(doc, allMatches, modeDescription);
        }

        private Result ApplyAutoMatchingResultsImpl(Document doc,
            List<(ViewSheet sheet, List<(Viewport source, Viewport target, string reason)> matches)> allMatches,
            string modeDescription)
        {
            using (Transaction trans = new Transaction(doc, $"自動マッチング位置変更 ({modeDescription})"))
            {
                trans.Start();

                int totalSuccess = 0;
                int totalProcessed = 0;

                foreach (var (sheet, matches) in allMatches)
                {
                    foreach (var (source, target, reason) in matches)
                    {
                        totalProcessed++;
                        try
                        {
                            XYZ sourcePosition = source.GetBoxCenter();
                            target.SetBoxCenter(sourcePosition);
                            totalSuccess++;
                        }
                        catch
                        {
                            // エラーは無視して続行
                        }
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        private Result ExecuteDirectViewportPasteImpl(Document doc, ViewSheet activeSheet, List<Viewport> viewports)
        {
            return ApplyPositionToViewportsImpl(doc, viewports, "直接選択ビューポート", null);
        }

        private Result ExecuteMultiSheetPasteImpl(Document doc, List<ViewSheet> sheets)
        {
            var sheetInfo = new List<(ViewSheet sheet, List<Viewport> viewports)>();
            int totalViewports = 0;

            foreach (var sheet in sheets)
            {
                var viewports = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .WhereElementIsNotElementType()
                    .Cast<Viewport>()
                    .Where(vp => vp.SheetId == sheet.Id)
                    .ToList();

                sheetInfo.Add((sheet, viewports));
                totalViewports += viewports.Count;
            }

            if (totalViewports == 0)
            {
                return Result.Failed;
            }

            using (Transaction trans = new Transaction(doc, "シート ビューポート位置変更"))
            {
                trans.Start();

                int successCount = 0;
                int totalProcessed = 0;

                foreach (var (sheet, viewports) in sheetInfo)
                {
                    foreach (Viewport viewport in viewports)
                    {
                        totalProcessed++;
                        try
                        {
                            viewport.SetBoxCenter(ViewportPositionClipboard.CopiedPosition);
                            successCount++;
                        }
                        catch
                        {
                            // エラーは無視して続行
                        }
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        private Result ExecuteViewBasedPasteImpl(Document doc, List<View> selectedViews)
        {
            var viewportInfo = new List<(View view, Viewport viewport, ViewSheet sheet)>();

            foreach (var view in selectedViews)
            {
                var viewport = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .WhereElementIsNotElementType()
                    .Cast<Viewport>()
                    .FirstOrDefault(vp => vp.ViewId == view.Id);

                if (viewport != null)
                {
                    ViewSheet sheet = doc.GetElement(viewport.SheetId) as ViewSheet;
                    viewportInfo.Add((view, viewport, sheet));
                }
            }

            if (viewportInfo.Count == 0)
            {
                return Result.Failed;
            }

            // ビューポートのリストを抽出して処理
            var viewports = viewportInfo.Select(vi => vi.viewport).ToList();
            return ApplyPositionToViewportsImpl(doc, viewports, "選択ビュー", viewportInfo);
        }

        private Result ApplyPositionToViewportsImpl(Document doc, List<Viewport> viewports, string operationType,
            List<(View view, Viewport viewport, ViewSheet sheet)> detailInfo = null)
        {
            // ビューポートに位置を適用する共通処理
            using (Transaction trans = new Transaction(doc, $"{operationType} ビューポート位置変更"))
            {
                trans.Start();

                int successCount = 0;
                int errorCount = 0;

                foreach (Viewport viewport in viewports)
                {
                    try
                    {
                        viewport.SetBoxCenter(ViewportPositionClipboard.CopiedPosition);
                        successCount++;
                    }
                    catch
                    {
                        errorCount++;
                    }
                }

                trans.Commit();

                // エラーがある場合のみメッセージ設定は呼び出し元で処理
            }

            return Result.Succeeded;
        }
    }
}