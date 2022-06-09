using ExcelValidator.Models;

using Newtonsoft.Json;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System.Diagnostics;

var excelRule = new ExcelRule
{
    ExcelWorksheetRules = new List<ExcelWorksheetRule>
    {
        new ExcelWorksheetRule
        {
            WorksheetName = "Planilha 1",
            ValidateHeaderNames = true,
            ExcelColumnRules = new List<ExcelColumnRule>
            {
                new ExcelColumnRule
                {
                    HeaderName = "Coluna 1",
                    ValueType = "string"
                },
                new ExcelColumnRule
                {
                    HeaderName = "Coluna 2",
                    ValueType = "number"
                },
                new ExcelColumnRule
                {
                    HeaderName = "Coluna 3",
                    ValueType = "date",
                    Format = "dd/MM/yyyy"
                },
                new ExcelColumnRule
                {
                    HeaderName = "Coluna 4",
                    ValueType = "datetime",
                    Format = "dd/MM/yyyy HH:mm:ss"
                },
                new ExcelColumnRule
                {
                    HeaderName = "Coluna 5",
                    ValueType = "time",
                    Format = "HH:mm:ss"
                },
                new ExcelColumnRule
                {
                    HeaderName = "Coluna 6",
                    ValueType = "guid"
                }
            }
        }
    }
};

var jsonData = JsonConvert.SerializeObject(excelRule, Formatting.Indented);

File.WriteAllText("excel_rules.json", jsonData);

var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel.xlsx");

var sw = Stopwatch.StartNew();

Console.WriteLine("Starting...");

var worksheet = new XSSFWorkbook(filePath);

var numOfSheets = worksheet.NumberOfSheets;

if (numOfSheets != excelRule.ExcelWorksheetRules.Count)
{
    throw new Exception("O número de abas é diferente do esperado");
}

for (int sheetNum = 0; sheetNum < worksheet.NumberOfSheets; sheetNum++)
{
    var sheet = worksheet.GetSheetAt(sheetNum);

    var sheetRule = excelRule.ExcelWorksheetRules[sheetNum];

    if (sheet.SheetName != sheetRule.WorksheetName)
    {
        throw new Exception($"O nome da aba era para ser '{sheetRule.WorksheetName}', porém foi encontrado '{sheet.SheetName}'");
    }

    for (int rowNum = 0; rowNum <= sheet.LastRowNum; rowNum++)
    {
        var row = sheet.GetRow(rowNum);

        if (row.LastCellNum != sheetRule.ExcelColumnRules.Count)
        {
            throw new Exception($"A linha {rowNum + 1} contém {row.LastCellNum} colunas, porém o esperado era de {sheetRule.ExcelColumnRules.Count} colunas");
        }

        for (int cellNum = 0; cellNum < row.LastCellNum; cellNum++)
        {
            var cell = row.GetCell(cellNum);

            if (rowNum == 0 && sheetRule.ValidateHeaderNames)
            {
                if (cell.StringCellValue != sheetRule.ExcelColumnRules[cellNum].HeaderName)
                {
                    throw new Exception($"O nome da coluna {cellNum + 1} era para ser '{sheetRule.ExcelColumnRules[cellNum].HeaderName}', porém foi encontrado '{cell.StringCellValue}'");
                }

                continue;
            }

            if (cell.CellType != CellType.Blank)
            {
                var columnRule = sheetRule.ExcelColumnRules[cellNum];

                if (columnRule.ValueType == "string")
                {
                    if (cell.CellType != CellType.String)
                    {
                        throw new Exception($"O valor da célula {cell.Address} era para ser do tipo '{columnRule.ValueType}', porém foi encontrado '{cell.CellType}'");
                    }
                }
                else if (columnRule.ValueType == "guid")
                {
                    if (cell.CellType != CellType.String
                        || !Guid.TryParse(cell.StringCellValue, out _))
                    {
                        throw new Exception($"O valor da célula {cell.Address} era para ser do tipo '{columnRule.ValueType}', porém foi encontrado '{cell.CellType}'");
                    }
                }
                else if (columnRule.ValueType == "number")
                {
                    if (cell.CellType != CellType.Numeric)
                    {
                        throw new Exception($"O valor da célula {cell.Address} era para ser do tipo '{columnRule.ValueType}', porém foi encontrado '{cell.CellType}'");
                    }
                }
                else if (columnRule.ValueType == "date"
                    || columnRule.ValueType == "datetime"
                    || columnRule.ValueType == "time")
                {
                    if (cell.CellType != CellType.Numeric
                        || !TryParseToOADate(cell.NumericCellValue))
                    {
                        throw new Exception($"O valor da célula {cell.Address} era para ser do tipo '{columnRule.ValueType}', porém foi encontrado '{cell.CellType}'");
                    }
                }
            }
        }
    }
}

bool TryParseToOADate(double oaNumber)
{
    try
    {
        var _ = DateTime.FromOADate(oaNumber);

        return true;
    }
    catch
    {
        return false;
    }
}

Console.WriteLine($"Done in {sw.ElapsedMilliseconds} ms");