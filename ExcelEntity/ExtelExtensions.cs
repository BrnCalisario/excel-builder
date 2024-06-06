namespace ExcelEntity;

using OfficeOpenXml;
using System.Linq;

public static class ExcelExtensions
{
    public static ExcelWorksheet GetWorksheet(this ExcelPackage package, string name)
    {
        return package.Workbook.Worksheets.First(ws => ws.Name == name);
    }

    public static void AddHeader(this ExcelWorksheet worksheet, string[] header, int offsetRow = 0, int offsetCol = 0)
    {
        for (int i = 1; i <= header.Length; i++)
        {
            worksheet.Cells[offsetRow + 1, offsetCol + i].Value = header[i - 1];
        }
    }

    public static void AddColumnData(this ExcelWorksheet worksheet, string[] data, int offsetRow = 0, int colIndex = 0)
    {
        for (int i = 1; i <= data.Length; i++)
        {
            worksheet.Cells[offsetRow + i, colIndex].Value = data[i - 1];
        }
    }

    public static ExcelWorksheetBuilder WithColumns(this ExcelWorksheetBuilder worksheet, string[][] data, int startColumn, int offsetRow = 0)
    {
        for (int i = 1; i <= data.Length; i++)
        {
            worksheet.Worksheet.AddColumnData(data[i - 1], offsetRow, startColumn + i);
        }

        return worksheet;

    }

    public static ExcelWorksheetBuilder WithHeader(this ExcelWorksheetBuilder worksheet, string[] header, int offsetRow = 0, int offsetCol = 0)
    {
        worksheet.Worksheet.AddHeader(header, offsetRow, offsetCol);
        return worksheet;
    }

}

public class ExcelWorksheetBuilder
{ 
    public ExcelWorksheet Worksheet { get; set; }
}