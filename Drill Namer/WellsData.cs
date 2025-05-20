using Newtonsoft.Json;
using System.Collections.Generic;

namespace Drill_Namer.Models
{
    public class WellsData
    {
        [JsonProperty("wells")]
        public List<DrillData> Wells { get; set; } = new List<DrillData>();
    }
}
