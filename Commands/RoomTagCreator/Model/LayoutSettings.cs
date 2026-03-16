namespace Tools28.Commands.RoomTagCreator.Model
{
    public class LayoutSettings
    {
        public bool IsHorizontal { get; set; }
        public int Count { get; set; }
        public double SpacingMm { get; set; }

        public LayoutSettings()
        {
            IsHorizontal = false;
            Count = 5;
            SpacingMm = 2.0;
        }

        public double SpacingFeet
        {
            get { return SpacingMm / 304.8; }
        }
    }
}
