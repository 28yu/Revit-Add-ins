using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Tools28.Localization;

namespace Tools28.Commands.Room3DColor
{
    /// <summary>
    /// 部屋の立体（DirectShape）生成・専用3Dビュー作成・表示絞り込みを担当
    /// </summary>
    public static class RoomSolidGenerator
    {
        /// <summary>生成した DirectShape を識別するためのコメントマーカー</summary>
        public const string ShapeMarker = "Tools28_Room3DColor";

        // 再実行時に既存の要素を名前で探すため多言語化しない（CLAUDE.md の分類C）。
        // アドインの言語を切り替えても前回作ったビュー・ワークセット・凡例を見つけられるようにする
        public const string DefaultViewName = "部屋3D色分け";
        public const string DefaultWorksetName = "部屋3D色分け";
        public const string LegendViewName = "部屋3D色分け 凡例";

        /// <summary>
        /// 配置済み（面積を持つ）部屋を収集
        /// </summary>
        public static List<Room> CollectRooms(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                // カテゴリ指定だけでは型が保証されないため OfType を使う（Cast は例外になる）
                .OfType<Room>()
                .Where(r => r.Location != null && r.Area > 0)
                .ToList();
        }

        /// <summary>
        /// 体積計算が無効なら有効化する。呼び出しはトランザクション内で行うこと。
        /// </summary>
        /// <returns>無効から有効へ変更した場合 true</returns>
        public static bool EnsureVolumeComputation(Document doc)
        {
            AreaVolumeSettings settings = AreaVolumeSettings.GetAreaVolumeSettings(doc);
            if (settings == null)
                return false;

            if (!settings.ComputeVolumes)
            {
                settings.ComputeVolumes = true;
                return true;
            }
            return false;
        }

        /// <summary>
        /// ダイアログ用に部屋情報（名前・レベル・パラメータ値）を構築
        /// </summary>
        public static List<RoomEntry> BuildRoomEntries(Document doc, IEnumerable<Room> rooms)
        {
            var entries = new List<RoomEntry>();

            foreach (var room in rooms)
            {
                var entry = new RoomEntry
                {
                    Id = room.Id,
                    Name = GetRoomName(room),
                    LevelName = room.Level?.Name ?? RoomColorManager.NoValueLabel
                };

                foreach (Parameter p in room.Parameters)
                {
                    try
                    {
                        if (p == null || p.Definition == null)
                            continue;
                        // ElementId 型（他要素参照）は色分けキーに不向きなので除外
                        if (p.StorageType == StorageType.None ||
                            p.StorageType == StorageType.ElementId)
                            continue;

                        string name = p.Definition.Name;
                        if (string.IsNullOrEmpty(name) || entry.Parameters.ContainsKey(name))
                            continue;

                        string val = p.StorageType == StorageType.String
                            ? p.AsString()
                            : p.AsValueString();

                        if (!string.IsNullOrWhiteSpace(val))
                            entry.Parameters[name] = val;
                    }
                    catch { }
                }

                entries.Add(entry);
            }

            return entries;
        }

        /// <summary>
        /// 全部屋に共通して存在するパラメータ名の一覧（色分け基準の候補）
        /// </summary>
        public static List<string> GetCommonParameterNames(IEnumerable<RoomEntry> entries)
        {
            var list = entries.ToList();
            if (list.Count == 0)
                return new List<string>();

            var names = new HashSet<string>(list[0].Parameters.Keys);
            foreach (var e in list.Skip(1))
                names.IntersectWith(e.Parameters.Keys);

            return names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string GetRoomName(Room room)
        {
            string name = room.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString();
            if (string.IsNullOrEmpty(name))
                name = room.Name;
            return string.IsNullOrEmpty(name) ? RoomColorManager.NoValueLabel : name;
        }

        /// <summary>
        /// 既存の部屋色分け DirectShape を削除する。トランザクション内で呼ぶこと。
        /// </summary>
        public static int DeleteExistingShapes(Document doc)
        {
            var ids = new FilteredElementCollector(doc)
                .OfClass(typeof(DirectShape))
                .Cast<DirectShape>()
                .Where(ds => GetComment(ds) == ShapeMarker)
                .Select(ds => ds.Id)
                .ToList();

            int count = 0;
            foreach (var id in ids)
            {
                try { doc.Delete(id); count++; } catch { }
            }
            return count;
        }

        /// <summary>
        /// 部屋の閉じたソリッドから DirectShape（汎用モデル）を生成。トランザクション内で呼ぶこと。
        /// worksetId が指定されている場合は生成要素をそのワークセットに割り当てる。
        /// </summary>
        /// <returns>生成した DirectShape のId。生成できない場合は null。</returns>
        public static ElementId CreateRoomShape(Document doc, Room room, WorksetId worksetId)
        {
            GeometryElement shell;
            try
            {
                // ClosedShell プロパティは体積計算が有効なとき部屋の閉じたソリッドを返す
                shell = room.ClosedShell;
            }
            catch
            {
                return null;
            }
            if (shell == null)
                return null;

            var geometry = new List<GeometryObject>();
            foreach (GeometryObject obj in shell)
            {
                if (obj is Solid solid && solid.Volume > 0 && solid.Faces.Size > 0)
                    geometry.Add(solid);
            }
            if (geometry.Count == 0)
                return null;

            var catId = new ElementId(BuiltInCategory.OST_GenericModel);
            if (!DirectShape.IsValidCategoryId(catId, doc))
                return null;

            DirectShape ds = DirectShape.CreateElement(doc, catId);
            if (ds == null)
                return null;

            try
            {
                ds.SetShape(geometry);
            }
            catch
            {
                try { doc.Delete(ds.Id); } catch { }
                return null;
            }

            // マーカーを保存（再生成時の識別用）
            SetComment(ds, ShapeMarker);

            // マーク パラメータに対象部屋の要素IDを記載
            SetMark(ds, room.Id.IntValue().ToString());

            // 指定ワークセットへ割り当て
            if (worksetId != null)
            {
                try
                {
                    Parameter wp = ds.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                    if (wp != null && !wp.IsReadOnly)
                        wp.Set(worksetId.IntegerValue);
                }
                catch { }
            }

            return ds.Id;
        }

        /// <summary>
        /// 部屋色分け専用ワークセットを取得または作成する。トランザクション内で呼ぶこと。
        /// ワークシェアされていないモデルでは null を返す。
        /// 作成したワークセットは既定で全ビュー非表示に設定する。
        /// 同名のワークセットに部屋色分け以外の要素が入っている場合は、それを非表示にすると
        /// モデル全体で要素が見えなくなるため使わずに null を返す（nameInUse = true）。
        /// </summary>
        public static WorksetId EnsureRoomWorkset(Document doc, string worksetName, out bool nameInUse)
        {
            nameInUse = false;
            if (!doc.IsWorkshared)
                return null;

            Workset existing = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .FirstOrDefault(w => w.Name == worksetName);

            if (existing != null && HasOtherElements(doc, existing.Id))
            {
                nameInUse = true;
                return null;
            }

            WorksetId id = existing != null ? existing.Id : Workset.Create(doc, worksetName).Id;

            // 既定の表示を「非表示」にする（新規ワークセットは既存ビューもこの既定に従う）
            try
            {
                var visSettings = WorksetDefaultVisibilitySettings
                    .GetWorksetDefaultVisibilitySettings(doc);
                visSettings.SetWorksetVisibility(id, false);
            }
            catch { }

            return id;
        }

        // ワークセットに部屋色分けの DirectShape 以外の要素があるか
        private static bool HasOtherElements(Document doc, WorksetId worksetId)
        {
            return new FilteredElementCollector(doc)
                .WherePasses(new ElementWorksetFilter(worksetId))
                .WhereElementIsNotElementType()
                .Any(e => !(e is DirectShape ds && GetComment(ds) == ShapeMarker));
        }

        /// <summary>
        /// 指定ビューでのみワークセットを表示にする。トランザクション内で呼ぶこと。
        /// </summary>
        public static void ShowWorksetInView(View view, WorksetId worksetId)
        {
            if (view == null || worksetId == null)
                return;
            try { view.SetWorksetVisibility(worksetId, WorksetVisibility.Visible); }
            catch { }
        }

        /// <summary>
        /// ワークシェアされていないモデル向けの代替処理。
        /// 指定ビュー以外の全ビューで部屋ソリッドを要素単位で非表示にする。トランザクション内で呼ぶこと。
        /// </summary>
        public static void HideShapesInOtherViews(
            Document doc, ElementId colorViewId, ICollection<ElementId> shapeIds)
        {
            if (shapeIds == null || shapeIds.Count == 0)
                return;

            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                // シート・集計表・製図ビュー・凡例にはモデル要素が表示されないため対象外（処理時間の短縮）
                .Where(v => !v.IsTemplate && v.Id != colorViewId
                    && !(v is ViewSheet) && !(v is ViewSchedule) && !(v is ViewDrafting)
                    && v.ViewType != ViewType.Legend)
                .ToList();

            foreach (var v in views)
            {
                var hideable = new List<ElementId>();
                foreach (var id in shapeIds)
                {
                    try
                    {
                        Element e = doc.GetElement(id);
                        if (e != null && e.CanBeHidden(v))
                            hideable.Add(id);
                    }
                    catch { }
                }

                if (hideable.Count > 0)
                {
                    try { v.HideElements(hideable); } catch { }
                }
            }
        }

        /// <summary>
        /// 専用の等角3Dビューを作成する。トランザクション内で呼ぶこと。
        /// このツールで作った同名ビューがあれば削除して作り直す。
        /// 利用者が作った同名ビューは削除せず、新しいビューは別名（末尾に識別子）で作る。
        /// </summary>
        public static View3D CreateColorView(Document doc, string viewName)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate && v.Name == viewName);
            if (existing != null && IsColorView(doc, existing))
            {
                try { doc.Delete(existing.Id); } catch { }
            }

            ElementId vftId = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional)?.Id;

            if (vftId == null)
                throw new InvalidOperationException(Loc.S("Room3D.Err.No3DViewFamilyType"));

            View3D view = View3D.CreateIsometric(doc, vftId);

            try { view.Name = viewName; }
            catch { view.Name = viewName + "_" + Guid.NewGuid().ToString("N").Substring(0, 6); }

            // ビューテンプレートを解除して手動制御を可能にする
            try { view.ViewTemplateId = ElementId.InvalidElementId; } catch { }

            // サーフェス塗り分けが確実に見えるよう陰線処理表示にする
            try { view.DisplayStyle = DisplayStyle.HLR; } catch { }
            try { view.DetailLevel = ViewDetailLevel.Fine; } catch { }

            return view;
        }

        // このツールが作った色分けビューか（IsolateRoomShapes で一般モデル以外を非表示にしている）
        private static bool IsColorView(Document doc, View3D view)
        {
            try
            {
                return view.GetCategoryHidden(new ElementId(BuiltInCategory.OST_Walls))
                    && view.GetCategoryHidden(new ElementId(BuiltInCategory.OST_Floors))
                    && !view.GetCategoryHidden(new ElementId(BuiltInCategory.OST_GenericModel));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 指定ビューを「部屋ソリッドのみ」表示にする。トランザクション内で呼ぶこと。
        /// </summary>
        public static void IsolateRoomShapes(Document doc, View3D view, ICollection<ElementId> roomShapeIds)
        {
            var keepCatId = new ElementId(BuiltInCategory.OST_GenericModel);
            var keepSet = new HashSet<ElementId>(roomShapeIds);

            // 汎用モデル以外の全カテゴリを非表示
            foreach (Category cat in doc.Settings.Categories)
            {
                if (cat.CategoryType != CategoryType.Model &&
                    cat.CategoryType != CategoryType.Annotation)
                    continue;

                if (cat.Id.IntValue() == keepCatId.IntValue())
                    continue;

                try
                {
                    if (view.CanCategoryBeHidden(cat.Id))
                        view.SetCategoryHidden(cat.Id, true);
                }
                catch { }
            }

            // 汎用モデルのうち、今回生成した部屋ソリッド以外を要素単位で非表示
            var others = new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .WhereElementIsNotElementType()
                .Where(e => !keepSet.Contains(e.Id))
                .Select(e => e.Id)
                .ToList();

            if (others.Count > 0)
            {
                try { view.HideElements(others); } catch { }
            }
        }

        private static string GetComment(Element e)
        {
            try
            {
                return e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.AsString();
            }
            catch { return null; }
        }

        private static void SetComment(Element e, string value)
        {
            try
            {
                Parameter p = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                if (p != null && !p.IsReadOnly)
                    p.Set(value);
            }
            catch { }
        }

        private static void SetMark(Element e, string value)
        {
            try
            {
                Parameter p = e.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                if (p != null && !p.IsReadOnly)
                    p.Set(value);
            }
            catch { }
        }
    }
}
