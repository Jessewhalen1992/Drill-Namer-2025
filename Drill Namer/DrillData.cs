using Newtonsoft.Json;

namespace Drill_Namer.Models
{
    public class DrillData
    {
        [JsonProperty("gw_id")]
        public int WellId { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("file_#")]
        public string FileNumber { get; set; }

        [JsonProperty("final_survey_date")]
        public string FinalSurveyDate { get; set; }

        [JsonProperty("revision_#")]
        public int RevisionNumber { get; set; }

        [JsonProperty("as_drilled")]
        public bool AsDrilled { get; set; }

        [JsonProperty("lateral_len")]
        public double LateralLength { get; set; }

        [JsonProperty("coordinates")]
        public CoordinatesData Coordinates { get; set; }
    }
}
