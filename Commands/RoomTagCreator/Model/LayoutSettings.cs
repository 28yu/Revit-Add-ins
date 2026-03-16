namespace Tools28.Commands.RoomTagCreator.Model
{
    public class LayoutSettings
    {
        public int Count { get; set; }
        public double SpacingMm { get; set; }

        public LayoutSettings()
        {
            Count = 2;
            SpacingMm = 2.0;
        }

        public double SpacingFeet
        {
            get { return SpacingMm / 304.8; }
        }
    }
}
