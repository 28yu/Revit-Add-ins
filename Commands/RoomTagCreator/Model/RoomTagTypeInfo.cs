using Autodesk.Revit.DB;

namespace Tools28.Commands.RoomTagCreator.Model
{
    public class RoomTagTypeInfo
    {
        public ElementId Id { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string DisplayName { get; set; }

        public RoomTagTypeInfo(FamilySymbol symbol)
        {
            Id = symbol.Id;
            FamilyName = symbol.FamilyName;
            TypeName = symbol.Name;
            DisplayName = $"{FamilyName} : {TypeName}";
        }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}
