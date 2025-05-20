using Newtonsoft.Json;
using System.Collections.Generic;

namespace Drill_Namer.Models
{
    /// <summary>
    /// Represents the coordinates data for a drill.
    /// </summary>
    public class CoordinatesData
    {
        [JsonProperty("wellbore")]
        public List<string> Wellbore { get; set; }

        [JsonProperty("leg_#")]
        public List<int> LegNumbers { get; set; }

        [JsonProperty("order")]
        public List<int> Order { get; set; }

        [JsonProperty("latitude")]
        public List<string> Latitudes { get; set; }    // Changed from List<double> to List<string>

        [JsonProperty("longitude")]
        public List<string> Longitudes { get; set; }   // Changed from List<double> to List<string>
    }
}
