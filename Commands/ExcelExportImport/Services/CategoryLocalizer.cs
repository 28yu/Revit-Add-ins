using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// BuiltInCategory を Tools28 の言語設定に応じてローカライズする。
    /// マッピングに無いカテゴリは Revit が返す名前 (revitName) にフォールバックする。
    /// </summary>
    internal static class CategoryLocalizer
    {
        private static readonly Dictionary<BuiltInCategory, string> _keys = new Dictionary<BuiltInCategory, string>
        {
            // === 構造・建築モデル ===
            { BuiltInCategory.OST_Walls, "Category.Walls" },
            { BuiltInCategory.OST_Floors, "Category.Floors" },
            { BuiltInCategory.OST_Roofs, "Category.Roofs" },
            { BuiltInCategory.OST_Ceilings, "Category.Ceilings" },
            { BuiltInCategory.OST_Doors, "Category.Doors" },
            { BuiltInCategory.OST_Windows, "Category.Windows" },
            { BuiltInCategory.OST_Stairs, "Category.Stairs" },
            { BuiltInCategory.OST_Ramps, "Category.Ramps" },
            { BuiltInCategory.OST_Railings, "Category.Railings" },
            { BuiltInCategory.OST_Columns, "Category.Columns" },
            { BuiltInCategory.OST_StructuralColumns, "Category.StructuralColumns" },
            { BuiltInCategory.OST_StructuralFraming, "Category.StructuralFraming" },
            { BuiltInCategory.OST_StructuralFoundation, "Category.StructuralFoundation" },
            { BuiltInCategory.OST_StructuralTruss, "Category.StructuralTruss" },
            { BuiltInCategory.OST_Rebar, "Category.Rebar" },
            { BuiltInCategory.OST_StructuralStiffener, "Category.Stiffener" },
            { BuiltInCategory.OST_Rooms, "Category.Rooms" },
            { BuiltInCategory.OST_Areas, "Category.Areas" },
            { BuiltInCategory.OST_Furniture, "Category.Furniture" },
            { BuiltInCategory.OST_FurnitureSystems, "Category.FurnitureSystems" },
            { BuiltInCategory.OST_Casework, "Category.Casework" },
            { BuiltInCategory.OST_GenericModel, "Category.GenericModel" },
            { BuiltInCategory.OST_SpecialityEquipment, "Category.SpecialityEquipment" },
            { BuiltInCategory.OST_Site, "Category.Site" },
            { BuiltInCategory.OST_Topography, "Category.Topography" },
            { BuiltInCategory.OST_Planting, "Category.Planting" },
            { BuiltInCategory.OST_Parking, "Category.Parking" },
            { BuiltInCategory.OST_Entourage, "Category.Entourage" },
            { BuiltInCategory.OST_Mass, "Category.Mass" },
            { BuiltInCategory.OST_CurtainWallPanels, "Category.CurtainPanels" },
            { BuiltInCategory.OST_CurtainWallMullions, "Category.CurtainMullions" },

            // === MEP ===
            { BuiltInCategory.OST_DuctCurves, "Category.Ducts" },
            { BuiltInCategory.OST_DuctFitting, "Category.DuctFittings" },
            { BuiltInCategory.OST_DuctAccessory, "Category.DuctAccessories" },
            { BuiltInCategory.OST_DuctTerminal, "Category.AirTerminals" },
            { BuiltInCategory.OST_PipeCurves, "Category.Pipes" },
            { BuiltInCategory.OST_PipeFitting, "Category.PipeFittings" },
            { BuiltInCategory.OST_PipeAccessory, "Category.PipeAccessories" },
            { BuiltInCategory.OST_Conduit, "Category.Conduits" },
            { BuiltInCategory.OST_CableTray, "Category.CableTrays" },
            { BuiltInCategory.OST_ElectricalEquipment, "Category.ElectricalEquipment" },
            { BuiltInCategory.OST_ElectricalFixtures, "Category.ElectricalFixtures" },
            { BuiltInCategory.OST_LightingFixtures, "Category.LightingFixtures" },
            { BuiltInCategory.OST_LightingDevices, "Category.LightingDevices" },
            { BuiltInCategory.OST_MechanicalEquipment, "Category.MechanicalEquipment" },
            { BuiltInCategory.OST_PlumbingFixtures, "Category.PlumbingFixtures" },
            { BuiltInCategory.OST_Sprinklers, "Category.Sprinklers" },
            { BuiltInCategory.OST_FireAlarmDevices, "Category.FireAlarmDevices" },
            { BuiltInCategory.OST_DataDevices, "Category.DataDevices" },
            { BuiltInCategory.OST_CommunicationDevices, "Category.CommunicationDevices" },
            { BuiltInCategory.OST_SecurityDevices, "Category.SecurityDevices" },
            { BuiltInCategory.OST_NurseCallDevices, "Category.NurseCallDevices" },
            { BuiltInCategory.OST_TelephoneDevices, "Category.TelephoneDevices" },

            // === 注釈 ===
            { BuiltInCategory.OST_Grids, "Category.Grids" },
            { BuiltInCategory.OST_Levels, "Category.Levels" },
            { BuiltInCategory.OST_Sheets, "Category.Sheets" },
            { BuiltInCategory.OST_Views, "Category.Views" },
            { BuiltInCategory.OST_Viewports, "Category.Viewports" },
            { BuiltInCategory.OST_TextNotes, "Category.TextNotes" },
            { BuiltInCategory.OST_GenericAnnotation, "Category.GenericAnnotation" },
            { BuiltInCategory.OST_RevisionClouds, "Category.RevisionClouds" },
            { BuiltInCategory.OST_ScheduleGraphics, "Category.Schedules" },
        };

        /// <summary>
        /// BuiltInCategory に対応するローカライズ済みカテゴリ名を返す。
        /// マッピングに無い場合は revitName を返す。
        /// </summary>
        public static string GetLocalizedName(BuiltInCategory cat, string revitName)
        {
            if (_keys.TryGetValue(cat, out var key))
            {
                var localized = Loc.S(key);
                if (localized != key)
                    return localized;
            }
            return revitName;
        }
    }
}
