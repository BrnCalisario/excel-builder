using OfficeOpenXml;

namespace ExcelEntity;
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
        var worksheet = this.package.Workbook.Worksheets.Add(name);
        return new ExcelWorksheetBuilder() { Worksheet = worksheet };
    }

    public ExcelBuilder WithWorksheet(string[] names)
    {
        foreach (var name in names)
        {
            this.package.Workbook.Worksheets.Add(name);
        }

        return this;
    }

    public ExcelPackage Build()
    {
        return package;
    }
}