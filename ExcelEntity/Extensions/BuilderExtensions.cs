namespace ExcelEntity.Extensions;

using ExcelEntity.Builder;

public static class BuilderExtensions
{
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
