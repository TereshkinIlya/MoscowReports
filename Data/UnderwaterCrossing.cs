using Data.CrossingParts;

namespace Data
{
    public class UnderwaterCrossing
    {
        public string? Id { get; set; }
        public DateOnly? DateInspection { get; set; }
        public string? TypeOfSurvey { get; set; }
        public string? PositionMT { get; set; }
        public DeviationsPVP? DeviationsPVP { get; set; }
        public RivebedProcesses? RivebedProcesses { get; set; }
        public string? Character { get; set; }
        public Coordinates? Coordinates { get; set; }
        public DeviationsRivebed? DeviationsRivebed { get; set; }
        public string? RepairInfo { get; set; }
        public WaterFlowRate? WaterRate { get; set; }
        public MaxSpeedWaterFlow? MaxSpeeds { get; set; }
        public UnderwaterCrossing()
        {
            DeviationsPVP = new();
            RivebedProcesses = new();
            Coordinates = new();
            DeviationsRivebed = new();
            WaterRate = new();
            MaxSpeeds = new();
        }
    }
}
