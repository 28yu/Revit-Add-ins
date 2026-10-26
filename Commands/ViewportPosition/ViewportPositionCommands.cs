using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Tools28.Localization;

namespace Tools28.Commands.ViewportPosition
{
    // ビューポート位置情報を保持する静的クラス
    public static class ViewportPositionClipboard
    {
        public static XYZ CopiedPosition { get; set; }
        public static string SourceViewportInfo { get; set; }
        public static string SourceSheetName { get; set; }
        /// <summary>コピー元の「図の位置」。ペースト時にこれとペースト先の図の位置を合わせる。</summary>
        public static ViewportAnchorInfo SourceAnchor { get; set; }
        public static bool HasCopiedPosition => CopiedPosition != null;

        public static void Clear()
        {
            CopiedPosition = null;
            SourceAnchor = null;
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
                    message = Loc.S("ViewportPos.NeedSheet");
                    return Result.Failed;
                }

                // 選択されたビューポートを取得
                var selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 0)
                {
                    message = Loc.S("ViewportPos.NeedViewport");
                    return Result.Failed;
                }

                // 選択された要素からビューポートを抽出
                var viewports = selectedIds
                    .Select(id => doc.GetElement(id))
                    .OfType<Viewport>()
                    .ToList();

                if (viewports.Count == 0)
                {
                    message = Loc.S("ViewportPos.NoViewportInSelection");
                    return Result.Failed;
                }

                if (viewports.Count > 1)
                {
                    message = Loc.S("ViewportPos.MultipleViewports");
                    return Result.Failed;
                }

                Viewport selectedViewport = viewports[0];

                // 位置情報を取得（枠の中心と、図の位置の両方）
                XYZ position = selectedViewport.GetBoxCenter();
                ViewportAnchorInfo anchor = ViewportAnchor.Capture(doc, selectedViewport);

                // パースの 3D ビューは図の位置を計算できず、枠の中心でそろえても意味が無い
                if (anchor.IsPerspective)
                {
                    message = Loc.S("ViewportPos.CopyPerspective");
                    return Result.Failed;
                }

                // 関連するビュー情報を取得
                View relatedView = doc.GetElement(selectedViewport.ViewId) as View;
                string viewInfo = relatedView != null ? relatedView.Name : Loc.S("ViewportPos.UnknownView");

                // クリップボードに保存
                ViewportPositionClipboard.CopiedPosition = position;
                ViewportPositionClipboard.SourceViewportInfo = viewInfo;
                ViewportPositionClipboard.SourceSheetName = activeSheet.Name;
                ViewportPositionClipboard.SourceAnchor = anchor;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("ViewportPos.CopyError"), ex.Message);
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
                    message = Loc.S("ViewportPos.NothingCopied");
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
                            return ExecuteAutoMatchingImpl(doc, selectedSheets, ViewMatcher.MatchingMode.ExactMatch, Loc.S("ViewportPos.Mode.ExactMatch"));
                        }
                        else
                        {
                            return ExecuteSingleSheetPasteImpl(doc, selectedSheets[0], ref message);
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

                // パターン3: シートを開いて何も選択していない場合は、開いているシートが対象（マニュアルの手順）
                if (activeSheet != null && selectedIds.Count == 0)
                {
                    return ExecuteSingleSheetPasteImpl(doc, activeSheet, ref message);
                }

                // どれでもない場合はエラー
                message = Loc.S("ViewportPos.NeedTarget");
                return Result.Failed;
            }
            catch (Exception ex)
            {
                message = string.Format(Loc.S("ViewportPos.PasteError"), ex.Message);
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
            // 対応付いたビューポートどうしで、図の位置を合わせる
            var moves = allMatches
                .SelectMany(m => m.matches)
                .Select(m => (ViewportAnchor.Capture(doc, m.source), m.target))
                .ToList();
            return MoveViewports(doc, moves,
                string.Format(Loc.S("ViewportPos.Txn.AutoMatch"), modeDescription));
        }

        private Result ExecuteDirectViewportPasteImpl(Document doc, ViewSheet activeSheet, List<Viewport> viewports)
        {
            return ApplyPositionToViewportsImpl(doc, viewports, Loc.S("ViewportPos.Op.DirectViewport"), null);
        }

        // シート1枚が対象のとき。以前はシート上の全ビューポートを同じ位置へ動かしていたため、
        // 複数あると全部が1か所に重なった。動かすのは「どれか特定できる1つ」だけにする。
        private Result ExecuteSingleSheetPasteImpl(Document doc, ViewSheet sheet, ref string message)
        {
            var viewports = new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .WhereElementIsNotElementType()
                .Cast<Viewport>()
                .Where(vp => vp.SheetId == sheet.Id)
                .ToList();

            if (viewports.Count == 0)
            {
                message = Loc.S("ViewportPos.NoViewportOnSheet");
                return Result.Failed;
            }

            // 1つだけならそれを動かす
            if (viewports.Count == 1)
            {
                return ApplyPositionToViewportsImpl(doc, viewports, Loc.S("ViewportPos.Op.Sheet"), null);
            }

            // 複数あるときは、コピー元とビュー名が同じもの（凡例など複数シートに置けるビュー）が1つに決まる場合だけ動かす
            string sourceName = ViewportPositionClipboard.SourceViewportInfo;
            var sameName = viewports
                .Where(vp => (doc.GetElement(vp.ViewId) as View)?.Name is string name
                             && string.Equals(name, sourceName, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (sameName.Count == 1)
            {
                return ApplyPositionToViewportsImpl(doc, sameName, Loc.S("ViewportPos.Op.Sheet"), null);
            }

            // どれを動かすか決められないので何も動かさず、ビューポートを選ぶよう案内する
            message = Loc.S("ViewportPos.SheetHasMultiple");
            return Result.Failed;
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
            return ApplyPositionToViewportsImpl(doc, viewports, Loc.S("ViewportPos.Op.SelectedView"), viewportInfo);
        }

        private Result ApplyPositionToViewportsImpl(Document doc, List<Viewport> viewports, string operationType,
            List<(View view, Viewport viewport, ViewSheet sheet)> detailInfo = null)
        {
            // コピー元の図の位置（旧版でコピーした場合などに無ければ、枠の中心だけ持たせる）
            var source = ViewportPositionClipboard.SourceAnchor
                ?? new ViewportAnchorInfo { BoxCenter = ViewportPositionClipboard.CopiedPosition };
            var moves = viewports.Select(vp => (source, vp)).ToList();
            return MoveViewports(doc, moves,
                string.Format(Loc.S("ViewportPos.Txn.ApplyOp"), operationType));
        }

        /// <summary>
        /// ペースト先のビューポートを、コピー元と図の位置がそろうように動かす共通処理。
        /// 図の位置を計算できないもの（Revit 2021・製図ビュー・凡例など）は枠の中心でそろえる。
        /// そろえられないもの（パース・縮尺違い・見る向き違い）は動かさない。
        /// 結果は、問題があったときだけ画面で知らせる（黙って失敗しない）。
        /// </summary>
        private Result MoveViewports(Document doc,
            List<(ViewportAnchorInfo source, Viewport target)> moves, string transactionName)
        {
            int moved = 0;
            int fallback = 0;
            var problems = new List<string>();

            using (Transaction trans = new Transaction(doc, transactionName))
            {
                trans.Start();

                foreach (var (source, target) in moves)
                {
                    var dst = ViewportAnchor.Capture(doc, target);
                    string name = dst.ViewName ?? Loc.S("ViewportPos.UnknownView");
                    try
                    {
                        if (source.IsPerspective || dst.IsPerspective)
                        {
                            problems.Add(string.Format(Loc.S("ViewportPos.Report.Item"), name,
                                Loc.S("ViewportPos.Reason.Perspective")));
                            continue;
                        }

                        if (source.OriginOnSheet != null && dst.OriginOnSheet != null)
                        {
                            string why = ViewportAnchor.Incompatibility(source, dst);
                            if (why != null)
                            {
                                problems.Add(string.Format(Loc.S("ViewportPos.Report.Item"), name, why));
                                continue;
                            }
                            // 図の位置の差だけ枠ごと動かす
                            XYZ delta = source.OriginOnSheet - dst.OriginOnSheet;
                            target.SetBoxCenter(target.GetBoxCenter() + new XYZ(delta.X, delta.Y, 0));
                        }
                        else
                        {
                            if (source.BoxCenter == null) continue;
                            target.SetBoxCenter(source.BoxCenter);
                            fallback++;
                        }
                        moved++;
                    }
                    catch (Exception ex)
                    {
                        DiagLog.Write($"[ViewportPos] 移動に失敗 view={name}: {ex.Message}");
                        problems.Add(string.Format(Loc.S("ViewportPos.Report.Failed"), name, ex.Message));
                    }
                }

                // 何も動かなかったときは、空の操作を「元に戻す」履歴に残さない
                if (moved > 0) trans.Commit();
                else trans.RollBack();
            }

            if (fallback > 0 || problems.Count > 0)
            {
                var lines = new List<string>
                {
                    string.Format(Loc.S("ViewportPos.Report.Moved"), moved - fallback),
                };
                if (fallback > 0)
                    lines.Add(string.Format(Loc.S("ViewportPos.Report.Fallback"), fallback));
                if (problems.Count > 0)
                {
                    lines.Add("");
                    lines.Add(Loc.S("ViewportPos.Report.SkippedHeader"));
                    lines.AddRange(problems.Take(20));
                    if (problems.Count > 20)
                        lines.Add(string.Format(Loc.S("Import.Done.More"), problems.Count - 20));
                }
                TaskDialog.Show(Loc.S("ViewportPos.Report.Title"), string.Join("\n", lines));
            }

            return Result.Succeeded;
        }
    }
}