using System.Linq;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>Saving as the other Excel format carries more than the cells.</summary>
    public class FormatConversionTests : TestBase
    {
        static WorkBook Source(ExcelFileFormat format)
        {
            var wb = WorkBook.Create(format);
            var ws = wb.DefaultWorkSheet; ws.Name = "Data";
            ws["A1"].Value = "x"; ws["A2"].Value = 2; ws["A3"].Formula = "=A2*A2";
            ws.GetCellAt(0, 1).AddComment("remember", "me", false);
            ws["A2:A3"].SaveAsNamedRange("Pair", true);
            wb.Metadata.Author = "tester"; wb.Metadata.Title = "Converted";
            wb.Metadata.CustomProperties.AddProperty("k", "v");
            wb.Metadata.CustomProperties.AddProperty("n", 7);
            return wb;
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX, ".xls")]
        [InlineData(ExcelFileFormat.XLS, ".xlsx")]
        public void SaveAs_OtherFormat_KeepsNamesPropertiesAndComments(ExcelFileFormat from, string extension)
        {
            Source(from).SaveAs(Tmp("converted" + extension));
            var back = WorkBook.Load(Tmp("converted" + extension));
            var ws = back.DefaultWorkSheet;

            Assert.Equal("x", ws["A1"].Value);
            Assert.Equal(4d, ws["A3"].Value);
            Assert.Contains("Pair", ws.GetNamedRanges());
            Assert.Contains("Data", ws.FindNamedRange("Pair"));
            Assert.Equal("tester", back.Metadata.Author);
            Assert.Equal("Converted", back.Metadata.Title);
            Assert.Equal("v", back.Metadata.CustomProperties.GetValue<string>("k"));
            Assert.Equal(7, back.Metadata.CustomProperties.GetValue<int>("n"));
            Assert.Equal("remember", ws.GetCellAt(0, 1).Comment?.Content);
        }
    }
}
