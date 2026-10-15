using System;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Xunit;

namespace ExcelImporter.Tests
{
    public class PersonModel
    {
        [ExcelColumn("Full Name")]
        public string Name { get; set; } = string.Empty;

        [ExcelColumn("Age")]
        public int Age { get; set; }

        [ExcelColumn("Email Address")]
        public string Email { get; set; } = string.Empty;

        [ExcelColumn(Index = 4)]
        public DateTime BirthDate { get; set; }

        [ExcelColumn(Ignore = true)]
        public string IgnoredField { get; set; } = string.Empty;
    }

    public class ExcelImporterTests
    {
        private readonly ExcelImporter _importer;

        public ExcelImporterTests()
        {
            _importer = new ExcelImporter();
        }

        private MemoryStream CreateSampleExcelStream()
        {
            var ms = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("People");

                // Headers
                worksheet.Cell(1, 1).Value = "Full Name";
                worksheet.Cell(1, 2).Value = "Age";
                worksheet.Cell(1, 3).Value = "Email Address";
                worksheet.Cell(1, 4).Value = "Birth Date";

                // Data row 1
                worksheet.Cell(2, 1).Value = "Alice Smith";
                worksheet.Cell(2, 2).Value = 30;
                worksheet.Cell(2, 3).Value = "alice@example.com";
                worksheet.Cell(2, 4).Value = new DateTime(1994, 5, 12);

                // Data row 2
                worksheet.Cell(3, 1).Value = "Bob Jones";
                worksheet.Cell(3, 2).Value = 25;
                worksheet.Cell(3, 3).Value = "bob@example.com";
                worksheet.Cell(3, 4).Value = new DateTime(1999, 10, 20);

                workbook.SaveAs(ms);
            }
            ms.Position = 0;
            return ms;
        }

        [Fact]
        public void ReadToDataTable_ValidExcelStream_ReturnsDataTable()
        {
            using var stream = CreateSampleExcelStream();
            var table = _importer.ReadToDataTable(stream);

            Assert.NotNull(table);
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(4, table.Columns.Count);
            Assert.Equal("Full Name", table.Columns[0].ColumnName);
            Assert.Equal("Alice Smith", table.Rows[0]["Full Name"].ToString());
            Assert.Equal("Bob Jones", table.Rows[1]["Full Name"].ToString());
        }

        [Fact]
        public void ReadToList_ValidExcelStream_ReturnsMappedList()
        {
            using var stream = CreateSampleExcelStream();
            var list = _importer.ReadToList<PersonModel>(stream);

            Assert.NotNull(list);
            Assert.Equal(2, list.Count);

            var alice = list.First(p => p.Name == "Alice Smith");
            Assert.Equal(30, alice.Age);
            Assert.Equal("alice@example.com", alice.Email);
            Assert.Equal(new DateTime(1994, 5, 12), alice.BirthDate);
            Assert.Equal(string.Empty, alice.IgnoredField);

            var bob = list.First(p => p.Name == "Bob Jones");
            Assert.Equal(25, bob.Age);
            Assert.Equal("bob@example.com", bob.Email);
            Assert.Equal(new DateTime(1999, 10, 20), bob.BirthDate);
        }

        [Fact]
        public void ReadToDataTable_CustomSheetName_ReadsSpecifiedSheet()
        {
            using var ms = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var sheet1 = workbook.Worksheets.Add("SheetOne");
                sheet1.Cell(1, 1).Value = "ColA";
                sheet1.Cell(2, 1).Value = "Val1";

                var sheet2 = workbook.Worksheets.Add("CustomSheet");
                sheet2.Cell(1, 1).Value = "TargetCol";
                sheet2.Cell(2, 1).Value = "TargetVal";

                workbook.SaveAs(ms);
            }
            ms.Position = 0;

            var options = new ExcelImportOptions { SheetName = "CustomSheet" };
            var dt = _importer.ReadToDataTable(ms, options);

            Assert.Single(dt.Rows);
            Assert.Equal("TargetCol", dt.Columns[0].ColumnName);
            Assert.Equal("TargetVal", dt.Rows[0][0].ToString());
        }

        [Fact]
        public void ReadToDataTable_FileNotFound_ThrowsFileNotFoundException()
        {
            Assert.Throws<FileNotFoundException>(() => _importer.ReadToDataTable("non_existent_file.xlsx"));
        }
    }
}
