namespace UnitTest;

using ExcelEntity;
using OfficeOpenXml;
using System;
using System.IO;

public class ExcelTests
{
    private static readonly string projectPath = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName;

    private static readonly string outputFolder = Path.Combine(projectPath, "output");

    public ExcelTests()
    {
        if(!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }
    }

    [Fact]
    public void CreateWithWorksheets()
    {
        var builder = new ExcelBuilder()
            .WithLicense(LicenseContext.NonCommercial);

        string[][] data = [
            ["John", "James", "Jonathan"], 
            ["18/12/2002", "19/03/2000"]
        ];

        builder
            .WithWorksheet("My Data")
            .WithHeader(["Name", "Birthdate"])
            .WithColumns(data, 0, 1);

        var excel = builder.Build();

        excel.SaveAs(Path.Combine(outputFolder, $"{nameof(CreateWithWorksheets)}.xlsx"));         
    }
}