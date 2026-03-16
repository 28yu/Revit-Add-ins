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
                .Where(r => r.Area > 0)
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
        /// ドキュメント内のビューテンプレート一覧を取得
        /// </summary>
        public static List<View> GetViewTemplates(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate)
                .OrderBy(v => v.Name)
                .ToList();
        }

        /// <summary>
        /// ViewFamilyType一覧を取得（平面図用）
        /// </summary>
        public static List<ViewFamilyType> GetViewFamilyTypes(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .Where(vft => vft.ViewFamily == ViewFamily.FloorPlan
                           || vft.ViewFamily == ViewFamily.CeilingPlan)
                .OrderBy(vft => vft.Name)
                .ToList();
        }

        /// <summary>
        /// 新規ビューを作成
        /// </summary>
        public static ViewPlan CreateNewView(Document doc, View sourceView, string viewName,
            ElementId viewFamilyTypeId, ElementId viewTemplateId)
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

            if (viewFamilyTypeId == null || viewFamilyTypeId == ElementId.InvalidElementId)
            {
                throw new InvalidOperationException("ビューファミリタイプが選択されていません。");
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

            // ユーザー指定のビューテンプレートを適用
            if (viewTemplateId != null && viewTemplateId != ElementId.InvalidElementId)
            {
                newView.ViewTemplateId = viewTemplateId;
            }

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

            int columnCount = 0;
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
                    // 引き出し線を有効化
                    newTag.HasLeader = true;
                    createdTags.Add(newTag);

                    // 最初のタグからサイズを取得
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
                    if (settings.IsHorizontal)
                    {
                        // 横並び: X軸右方向、指定列数で改行
                        columnCount++;
                        if (columnCount % settings.Count == 0)
                        {
                            currentU = startU;
                            currentV -= height;
                        }
                        else
                        {
                            currentU += width;
                        }
                    }
                    else
                    {
                        // 縦並び: Y軸下方向、指定行数で右へシフト
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

            foreach (Category cat in categories)
            {
                if (cat.CategoryType != CategoryType.Model &&
                    cat.CategoryType != CategoryType.Annotation)
                    continue;

                // 部屋タグ (OST_RoomTags) は表示を維持
                if (cat.Id.IntegerValue == (int)BuiltInCategory.OST_RoomTags)
                    continue;

                // 部屋 (OST_Rooms) は表示を維持し、色塗り潰し以外のサブカテゴリを非表示
                if (cat.Id.IntegerValue == (int)BuiltInCategory.OST_Rooms)
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
        }
    }
}
