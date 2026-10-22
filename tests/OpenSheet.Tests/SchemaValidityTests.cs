using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenSheet.DataValidations;
using OpenSheet.Drawing.Charts;
using OpenSheet.Drawing.Images;
using OpenSheet.Styles;
using Xunit;
using Xunit.Abstractions;
using CfOp = OpenSheet.Formatting.Enums.ComparisonOperator;
using DvOp = OpenSheet.DataValidations.ComparisonOperator;

namespace OpenSheet.Tests
{
    /// <summary>
    /// Validates what we write against the OOXML schema using Microsoft's own SDK validator (an independent reader,
    /// unlike NPOI which would happily re-read its own mistakes).
    /// </summary>
    public class SchemaValidityTests : TestBase
    {
        readonly ITestOutputHelper _out;
        public SchemaValidityTests(ITestOutputHelper output) { _out = output; }

        static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

        string[] Validate(string path)
        {
            using (var doc = SpreadsheetDocument.Open(path, false))
            {
                var v = new OpenXmlValidator(FileFormatVersions.Office2019);
                return v.Validate(doc).Select(e => $"{e.Description}  @ {e.Path?.XPath}  [{e.Part?.Uri}]").ToArray();
            }
        }

        void AssertValid(string path)
        {
            var errors = Validate(path);
            foreach (var e in errors) _out.WriteLine(e);
            Assert.True(errors.Length == 0, $"{errors.Length} schema error(s):\n" + string.Join("\n", errors.Take(15)));
        }

        [Fact]
        public void Basic_Values_Styles_Formulas_AreValid()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "text"; ws["A2"].Value = 1.5; ws["A3"].Value = true; ws["A4"].Value = DateTime.Today;
            ws["B1"].Formula = "=SUM(A2:A2)*2";
            ws["A1"].Style.Font.Bold = true; ws["A1"].Style.SetBackgroundColor("#FFCC00");
            ws["A1"].Style.TopBorder.SetBorderStyle("#000000", BorderType.Medium);
            ws["A2"].FormatString = "#,##0.00"; ws["A2"].Style.HorizontalAlignment = HorizontalAlignment.Right;
            ws["A1:C1"].Style.WrapText = true;
            wb.SaveAs(Tmp("basic.xlsx"));
            AssertValid(Tmp("basic.xlsx"));
        }

        [Fact]
        public void KitchenSink_IsValid()
        {
            Samples.KitchenSink().SaveAs(Tmp("sink.xlsx"));
            AssertValid(Tmp("sink.xlsx"));
        }

        [Fact]
        public void DiagonalDirections_ChangedOperators_ReplacedLinks_AndRemovedLoadedCharts_AreValid()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            for (var i = 0; i < 4; i++) { ws.GetCellAt(i, 0).Value = "k" + i; ws.GetCellAt(i, 1).Value = i + 1; }
            ws["D1:F1"].Style.DiagonalBorder.SetBorderStyle("#FF0000", BorderType.Thin);
            ws["D1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Both;
            ws["E1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;
            ws["G1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Backward;
            ws["H1"].Style.FillPattern = FillPattern.SolidForeground;
            ws["H1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.None;
            var rule = ws.DataValidations.AddIntegerRule("C1:C3", DvOp.Between, 1, 9);
            rule.ComparisonOperator = DvOp.NotBetween;
            ws.GetCellAt(0, 0).Hyperlink = "https://example.com/one";
            ws.GetCellAt(0, 0).Hyperlink = "mailto:me@example.com";
            var keep = ws.CreateChart(ChartType.Column, 6, 0, 16, 5); keep.AddSeries("A1:A4", "B1:B4"); keep.Plot();
            var drop = ws.CreateChart(ChartType.Line, 18, 0, 28, 5); drop.AddSeries("A1:A4", "B1:B4"); drop.Plot();
            wb.SaveAs(Tmp("first.xlsx"));
            AssertValid(Tmp("first.xlsx"));

            var again = WorkBook.Load(Tmp("first.xlsx"));
            var sheet = again.DefaultWorkSheet;
            Assert.Equal(2, sheet.Charts.Count);
            sheet.RemoveChart(sheet.Charts.Last());
            again.WorkSheets[0] = Samples.KitchenSink().DefaultWorkSheet;   // a whole sheet replaced by one from another workbook
            again.SaveAs(Tmp("second.xlsx"));
            AssertValid(Tmp("second.xlsx"));

            var third = WorkBook.Load(Tmp("first.xlsx"));
            third.DefaultWorkSheet.RemoveChart(third.DefaultWorkSheet.Charts.Last());
            third.SaveAs(Tmp("third.xlsx"));
            AssertValid(Tmp("third.xlsx"));
            Assert.Single(WorkBook.Load(Tmp("third.xlsx")).DefaultWorkSheet.Charts);
        }

        [Fact]
        public void Reloaded_AndResaved_StaysValid()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "x"; ws["A2"].Formula = "=1+1";
            ws.DataValidations.AddIntegerRule("B1", DvOp.Equal, 3, null);
            ws.InsertImage(Png, ImageFormat.PNG, 3, 3, 5, 5);
            wb.SaveAs(Tmp("one.xlsx"));
            var again = WorkBook.Load(Tmp("one.xlsx"));
            again.DefaultWorkSheet["A3"].Value = "edited";
            again.SaveAs(Tmp("two.xlsx"));
            AssertValid(Tmp("two.xlsx"));
        }
    }
}
