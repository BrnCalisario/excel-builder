namespace ExcelEntity.Extensions;

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
}