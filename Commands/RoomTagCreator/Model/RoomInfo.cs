using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace Tools28.Commands.RoomTagCreator.Model
{
    public class RoomInfo
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public Room Element { get; set; }

        public RoomInfo(Room room)
        {
            Element = room;
            Id = room.Id;
            Name = room.LookupParameter("名前")?.AsString() ?? room.Name ?? "(名称なし)";
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
