using Newtonsoft.Json;

namespace ExcelValidator.Models
{
    public class ExcelWorksheetRule
    {
        [JsonProperty("worksheet_name")]
        public string WorksheetName { get; set; }

        [JsonProperty("validate_header_names")]
        public bool ValidateHeaderNames { get; set; }

        [JsonProperty("excel_column_rules")]
        public List<ExcelColumnRule> ExcelColumnRules { get; set; }
    }
}