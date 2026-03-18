using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Tools28.Commands.RoomTagCreator.Model;

namespace Tools28.Commands.RoomTagCreator
{
    public static class RoomTagService
    {
        /// <summary>
        /// ビューポート内の部屋を取得
        /// </summary>
        public static List<RoomInfo> GetViewportRooms(Document doc, Viewport viewport)
        {
            ElementId viewId = viewport.ViewId;
            View view = doc.GetElement(viewId) as View;
            if (view == null)
                return new List<RoomInfo>();

            var rooms = new FilteredElementCollector(doc, viewId)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r => r.Location != null)
                .OrderBy(r => r.LookupParameter("名前")?.AsString() ?? r.Name ?? "")
                .Select(r => new RoomInfo(r))
                .ToList();

            return rooms;
        }

        /// <summary>
        /// ドキュメント内の全RoomTagタイプを取得
        /// </summary>
        public static List<RoomTagTypeInfo> GetRoomTagTypes(Document doc)
        {
            var tagTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RoomTags)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .OrderBy(s => s.FamilyName)
                .ThenBy(s => s.Name)
                .Select(s => new RoomTagTypeInfo(s))
                .ToList();

            return tagTypes;
        }

        /// <summary>
        /// 指定名のViewFamilyTypeを取得、なければ複製して作成
        /// </summary>
        private static ElementId GetOrCreateViewFamilyType(Document doc, string typeName)
        {
            // 天井伏図かどうかを判定
            bool isCeilingPlan = typeName.Contains("天井伏図");
            ViewFamily targetFamily = isCeilingPlan ? ViewFamily.CeilingPlan : ViewFamily.FloorPlan;

            var allTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .Where(vft => vft.ViewFamily == targetFamily)
                .ToList();

            // 同名が既にあればそれを使用
            var existing = allTypes.FirstOrDefault(vft => vft.Name == typeName);
            if (existing != null)
                return existing.Id;

            // なければ最初のタイプを複製
            if (allTypes.Count == 0)
                throw new InvalidOperationException($"{targetFamily} のビューファミリタイプが見つかりません。");

            var duplicated = allTypes.First().Duplicate(typeName);
            return duplicated.Id;
        }

        /// <summary>
        /// 新規ビューを作成
        /// </summary>
        public static ViewPlan CreateNewView(Document doc, View sourceView, string viewName,
            string viewFamilyTypeName)
        {
            // ソースビューのレベルを取得
            Level level = null;
            if (sourceView is ViewPlan viewPlan)
            {
                level = viewPlan.GenLevel;
            }

            if (level == null)
            {
                throw new InvalidOperationException("ソースビューのレベルが取得できません。");
            }

            // ビュー名の重複チェック
            var existingViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate)
                .Select(v => v.Name)
                .ToList();

            if (existingViews.Contains(viewName))
            {
                throw new InvalidOperationException($"ビュー名「{viewName}」は既に存在します。別の名前を指定してください。");
            }

            // ビューファミリタイプを取得または作成
            ElementId viewFamilyTypeId = GetOrCreateViewFamilyType(doc, viewFamilyTypeName);

            // 新規平面図を作成
            ViewPlan newView = ViewPlan.Create(doc, viewFamilyTypeId, level.Id);
            newView.Name = viewName;

            // クロップボックスをソースビューからコピー
            if (sourceView.CropBoxActive)
            {
                newView.CropBoxActive = true;
                newView.CropBox = sourceView.CropBox;
            }

            // クロップボックス表示を非表示
            newView.CropBoxVisible = false;

            // ビューテンプレートを解除（自動適用含む）して手動設定を可能にする
            newView.ViewTemplateId = ElementId.InvalidElementId;

            // 部屋タグ以外の全カテゴリを非表示にする
            HideAllCategoriesExceptRoomTags(doc, newView);

            return newView;
        }

        /// <summary>
        /// 部屋タグを自動配置
        /// </summary>
        public static List<RoomTag> CreateRoomTags(Document doc, ViewPlan view,
            List<RoomInfo> rooms, ElementId tagTypeId, LayoutSettings settings)
        {
            var createdTags = new List<RoomTag>();

            if (rooms == null || rooms.Count == 0)
                return createdTags;

            if (tagTypeId == null || tagTypeId == ElementId.InvalidElementId)
                return createdTags;

            int viewScale = view.Scale;

            // クロップボックスの左上を起点にする
            BoundingBoxXYZ cropBox = view.CropBox;
            double startU = cropBox.Min.X;
            double startV = cropBox.Max.Y;

            // 配置位置の計算用変数
            double width = 0;
            double height = 0;
            bool sizeCalculated = false;

            int rowCount = 0;
            double currentU = startU;
            double currentV = startV;

            for (int i = 0; i < rooms.Count; i++)
            {
                var room = rooms[i];
                UV tagPosition = new UV(currentU, currentV);

                // 部屋タグを作成
                RoomTag newTag = doc.Create.NewRoomTag(
                    new LinkElementId(room.Id), tagPosition, view.Id);

                if (newTag != null)
                {
                    // タグタイプを設定
                    newTag.ChangeTypeId(tagTypeId);
                    // 引き出し線は後で一括設定（BoundingBox取得時に影響するため）
                    newTag.HasLeader = false;
                    createdTags.Add(newTag);

                    // 最初のタグからサイズを取得（引き出し線なしの状態で測定）
                    if (!sizeCalculated)
                    {
                        doc.Regenerate();
                        BoundingBoxXYZ tagBB = newTag.get_BoundingBox(view);
                        if (tagBB != null)
                        {
                            // BoundingBoxはモデル座標（スケール反映済み）
                            // 間隔は紙面mm指定なのでスケール変換
                            double stepModel = settings.SpacingFeet * viewScale;
                            width = (tagBB.Max.X - tagBB.Min.X) + stepModel;
                            height = (tagBB.Max.Y - tagBB.Min.Y) + stepModel;
                            sizeCalculated = true;
                        }
                        else
                        {
                            // BoundingBox取得失敗時のデフォルト値（紙面10mm×5mm）
                            width = 10.0 / 304.8 * viewScale;
                            height = 5.0 / 304.8 * viewScale;
                            sizeCalculated = true;
                        }
                    }

                    // 次のタグの位置を計算
                    // Count = 行数。縦にCount個並べたら右にシフト
                    rowCount++;
                    if (rowCount % settings.Count == 0)
                    {
                        currentU += width;
                        currentV = startV;
                    }
                    else
                    {
                        currentV -= height;
                    }
                }
            }

            // 全タグの引き出し線を一括で有効化
            foreach (var tag in createdTags)
            {
                tag.HasLeader = true;
            }

            return createdTags;
        }

        /// <summary>
        /// ビュー名を自動生成
        /// パターン: 元ビュー名 "xxx_yyy - zzz" → "仕上表_yyy - zzz"
        /// </summary>
        public static string GenerateViewName(string sourceViewName)
        {
            if (string.IsNullOrEmpty(sourceViewName))
                return "仕上表";

            // パターン: xxx_yyy - zzz
            var match = System.Text.RegularExpressions.Regex.Match(
                sourceViewName, @"^(.*)_(.*)$");

            if (match.Success)
            {
                string remainder = match.Groups[2].Value;
                return $"仕上表_{remainder}";
            }

            return $"仕上表_{sourceViewName}";
        }

        /// <summary>
        /// 部屋タグ・部屋（色塗り潰しのみ）以外の全カテゴリを非表示にする
        /// </summary>
        private static void HideAllCategoriesExceptRoomTags(Document doc, View view)
        {
            Categories categories = doc.Settings.Categories;

            // 表示を維持するカテゴリ
#if REVIT2026
            var keepVisible = new HashSet<long>
            {
                (long)BuiltInCategory.OST_RoomTags,
                (long)BuiltInCategory.OST_RevisionClouds,       // 改定雲マーク
                (long)BuiltInCategory.OST_RevisionCloudTags,    // 改定雲マークタグ
                (long)BuiltInCategory.OST_TextNotes,            // 文字注記
            };
#else
            var keepVisible = new HashSet<int>
            {
                (int)BuiltInCategory.OST_RoomTags,
                (int)BuiltInCategory.OST_RevisionClouds,       // 改定雲マーク
                (int)BuiltInCategory.OST_RevisionCloudTags,    // 改定雲マークタグ
                (int)BuiltInCategory.OST_TextNotes,            // 文字注記
            };
#endif

            foreach (Category cat in categories)
            {
                if (cat.CategoryType != CategoryType.Model &&
                    cat.CategoryType != CategoryType.Annotation)
                    continue;

                // 表示を維持するカテゴリはスキップ
#if REVIT2026
                if (keepVisible.Contains(cat.Id.Value))
                    continue;
#else
                if (keepVisible.Contains(cat.Id.IntegerValue))
                    continue;
#endif

                // 部屋 (OST_Rooms) は表示を維持し、色塗り潰し以外のサブカテゴリを非表示
#if REVIT2026
                if (cat.Id.Value == (long)BuiltInCategory.OST_Rooms)
#else
                if (cat.Id.IntegerValue == (int)BuiltInCategory.OST_Rooms)
#endif
                {
                    foreach (Category subCat in cat.SubCategories)
                    {
                        // 色塗り潰し (Color Fill) は表示を維持
                        string subName = subCat.Name;
                        if (subName == "色塗り潰し" || subName == "Color Fill")
                            continue;

                        try
                        {
                            if (view.CanCategoryBeHidden(subCat.Id))
                            {
                                view.SetCategoryHidden(subCat.Id, true);
                            }
                        }
                        catch { }
                    }
                    continue;
                }

                // その他のカテゴリを非表示
                try
                {
                    if (view.CanCategoryBeHidden(cat.Id))
                    {
                        view.SetCategoryHidden(cat.Id, true);
                    }
                }
                catch { }

                // サブカテゴリも非表示
                foreach (Category subCat in cat.SubCategories)
                {
                    try
                    {
                        if (view.CanCategoryBeHidden(subCat.Id))
                        {
                            view.SetCategoryHidden(subCat.Id, true);
                        }
                    }
                    catch { }
                }
            }

            // 部屋タグの引き出し線を白色・間隔が広いドットに設定
            SetRoomTagLeaderOverride(doc, view);
        }

        /// <summary>
        /// 部屋タグの引き出し線を白色・間隔が広いドットにオーバーライド
        /// </summary>
        private static void SetRoomTagLeaderOverride(Document doc, View view)
        {
            // 間隔が広いドットのラインパターンを取得または作成
            ElementId linePatternId = GetOrCreateWideDotPattern(doc);

            var overrides = new OverrideGraphicSettings();
            overrides.SetProjectionLineColor(new Color(255, 255, 255));
            if (linePatternId != null && linePatternId != ElementId.InvalidElementId)
            {
                overrides.SetProjectionLinePatternId(linePatternId);
            }

            var roomTagCatId = new ElementId(BuiltInCategory.OST_RoomTags);
            view.SetCategoryOverrides(roomTagCatId, overrides);
        }

        /// <summary>
        /// 間隔が広いドットのラインパターンを取得または作成
        /// </summary>
        private static ElementId GetOrCreateWideDotPattern(Document doc)
        {
            string patternName = "間隔が広いドット";

            // 既存のパターンを検索（日本語名・英語名）
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(LinePatternElement))
                .Cast<LinePatternElement>()
                .FirstOrDefault(lp => lp.Name == patternName
                    || lp.Name == "Wide Dot"
                    || lp.Name == "Dot (wide)");

            if (existing != null)
                return existing.Id;

            // 見つからない場合は作成（ドット + 広い間隔）
            var segments = new List<LinePatternSegment>
            {
                new LinePatternSegment(LinePatternSegmentType.Dot, 0),
                new LinePatternSegment(LinePatternSegmentType.Space, 30.0 / 304.8)  // 30mm間隔
            };

            var pattern = new LinePattern(patternName);
            pattern.SetSegments(segments);

            return LinePatternElement.Create(doc, pattern).Id;
        }
    }
}
