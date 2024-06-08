namespace ExcelEntity.Builder;

using ExcelEntity.Extensions;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

public class ExcelBuilder
{
    private readonly ExcelPackage package = new();

    public ExcelBuilder WithLicense(LicenseContext license)
    {
        ExcelPackage.LicenseContext = license;
        return this;
    }

    public ExcelWorksheetBuilder WithWorksheet(string name)
    {
        var worksheet = package.Workbook.Worksheets.Add(name);
        return new ExcelWorksheetBuilder() { Worksheet = worksheet };
    }

    public ExcelBuilder WithWorksheet<T>(List<T> data, List<string> fields = null, string name = null)
    {
        name ??= typeof(T).Name;

        var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        if (fields != null)
        {
            properties = properties.Where(p => fields.Contains(p.Name)).ToArray();
        }

        var worksheet = package.Workbook.Worksheets.Add(name);

        worksheet.AddHeader(properties.Select(p => p.Name).ToArray());

        for (int i = 0; i < data.Count; i++) {
            for (int j = 0; j < properties.Length; j++) {
                worksheet.Cells[i + 2, j + 1].Value = properties[j].GetValue(data[i]).ToString();
            }
        }

        return this;
    }

    public ExcelBuilder WithWorksheet(string[] names)
    {
        foreach (var name in names)
        {
            package.Workbook.Worksheets.Add(name);
        }

        return this;
    }

    public ExcelPackage Build()
    {
        return package;
    }
}