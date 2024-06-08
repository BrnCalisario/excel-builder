namespace UnitTest;

using ExcelEntity.Builder;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
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

        var data = new List<Person>
        {
            new Person { Name = "John", BirthDate = new DateTime(2002, 12, 18) },
            new Person { Name = "James", BirthDate = new DateTime(2000, 3, 19) },
            new Person { Name = "Jonathan", BirthDate = new DateTime(1999, 5, 20) }
        };

        builder.WithWorksheet(data, new List<string> { "Name", "BirthDate" }, "People");
    
        var excel = builder.Build();

        excel.SaveAs(Path.Combine(outputFolder, $"{nameof(CreateWithWorksheets)}.xlsx"));         
    }
}