using Newtonsoft.Json;

namespace ExcelValidator.Models
{
    public class ExcelRule
    {
        [JsonProperty("excel_worksheet_rules")]
        public List<ExcelWorksheetRule> ExcelWorksheetRules { get; set; }
    }
}
