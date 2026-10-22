using System;
using Xunit;

namespace OpenSheet.Tests
{
    public class MetadataTests : TestBase
    {
        [Theory]
        [InlineData("m.xlsx")]
        [InlineData("m.xls")]
        public void Properties_RoundTrip(string file)
        {
            var wb = WorkBook.Create(file.EndsWith("xls") ? ExcelFileFormat.XLS : ExcelFileFormat.XLSX);
            var m = wb.Metadata;
            m.Author = "Ana"; m.Title = "Report"; m.Subject = "Q3"; m.Keywords = "a,b"; m.Comments = "hello";
            m.Category = "Finance";
            m.Created = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc);
            m.CustomProperties.AddProperty("Owner", "Bob");
            m.CustomProperties.AddProperty("Count", 7);
            m.CustomProperties.AddProperty("Ratio", 0.5);
            m.CustomProperties.AddProperty("Flag", true);
            wb.SaveAs(Tmp(file));

            var b = WorkBook.Load(Tmp(file)).Metadata;
            Assert.Equal("Ana", b.Author);
            Assert.Equal("Report", b.Title);
            Assert.Equal("Q3", b.Subject);
            Assert.Equal("a,b", b.Keywords);
            Assert.Equal("hello", b.Comments);
            Assert.Equal("Finance", b.Category);
            Assert.Equal(2024, b.Created.Value.Year);
            Assert.Equal("Bob", b.CustomProperties.GetValue<string>("Owner"));
            Assert.Equal(7, b.CustomProperties.GetValue<int>("Count"));
            Assert.Equal(0.5, b.CustomProperties.GetValue<double>("Ratio"));
            Assert.True(b.CustomProperties.GetValue<bool>("Flag"));
            Assert.True(b.CustomProperties.Contains("Owner"));
            Assert.False(b.CustomProperties.TryGetValue<int>("Missing", out _));
            Assert.True(b.CustomProperties.Remove("Owner"));
            Assert.False(b.CustomProperties.Contains("Owner"));
        }

        [Fact]
        public void CustomProperty_AddTwice_Replaces()
        {
            var cp = WorkBook.Create().Metadata.CustomProperties;
            cp.AddProperty("k", 1);
            cp.AddProperty("k", 2);
            Assert.Equal(2, cp.GetValue<int>("k"));
        }
    }
}
