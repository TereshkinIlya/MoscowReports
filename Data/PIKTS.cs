
namespace Data
{
    public class PIKTS
    {
        public string ID {  get; set; }
        public string? ConditionOfPipeline { get; set; }
        public string? SteelGrade { get; set; }
        public string? DateLastVTD { get; set; }
        public Dictionary<string, string?> DefectsPoima { get; set; }
        public Dictionary<string, string?> DefectsRuslo { get; set; }
        public string? SafePeriod { get; set; }
        public string? OTSNumber { get; set; }
        public string? DateReport { get; set; }
        public string? DateVTD { get; set; }
        public string? DateDefect { get; set; }
        public string? DateCDS { get; set; }
        public string? DateJumpers { get; set; }
        public string? DateLimited { get; set; }
        public string? DateVRK { get; set; }
        public string? DateUZA { get; set; }
        public string? DateWeldedElement { get; set; }
        public string? DateConnectedDetails { get; set; }
        public string? DateKPSOD { get; set; }
        public string? DateDrainageContainers { get; set; }
        public string? DatePVP { get; set; }
        public string? DateCorrosion { get; set; }
        public string? Organization { get; set; }
        public string? Events { get; set; }
        public PIKTS()
        {
            DefectsPoima = new Dictionary<string, string?>
            {
                {DateTime.Now.Year.ToString(),null},
                {(DateTime.Now.Year + 1).ToString(),null},
                {(DateTime.Now.Year + 2).ToString(),null}
            };
            DefectsRuslo = new Dictionary<string, string?>
            {
                {DateTime.Now.Year.ToString(),null},
                {(DateTime.Now.Year + 1).ToString(),null},
                {(DateTime.Now.Year + 2).ToString(),null}
            };
        }
    }
}
