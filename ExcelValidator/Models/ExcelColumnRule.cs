using Newtonsoft.Json;

namespace ExcelValidator.Models
{
    public class ExcelColumnRule
    {
        [JsonProperty("header_name")]
        public string HeaderName { get; set; }

        [JsonProperty("value_type")]
        public string ValueType { get; set; }

        [JsonProperty("format")]
        public string? Format { get; set; }
    }
}