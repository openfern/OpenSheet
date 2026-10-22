using System.Linq;
using DvOp = OpenSheet.DataValidations.ComparisonOperator;
using OpenSheet.Drawing.Charts;
using OpenSheet.Drawing.Images;
using OpenSheet.Styles;
using Xunit;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Features edited after they were created (a diagonal direction, a validation operator, a replaced link), and sheets removed or
    /// replaced, must read the same in Excel as in OpenSheet, and Excel must open the result.
    /// </summary>
    [Collection("Excel")]
    public class EditedFeaturesOracleTests : TestBase
    {
        readonly ExcelOracle _excel;
        public EditedFeaturesOracleTests(ExcelOracle excel) { _excel = excel; }

        const int None = -4142;   // xlNone

        [ExcelFact]
        public void DiagonalDirections_ValidationOperator_AndLinkEdits_ReadTheSameInExcel()
        {
            var wb = WorkBook.Create(); var ws = wb.DefaultWorkSheet; ws.Name = "S";
            ws["D1:F1"].Style.DiagonalBorder.SetBorderStyle("#FF0000", BorderType.Thin);
            ws["D1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Both;        // after the line exists
            ws["E1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;
            ws["H1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Backward;   // before the line exists
            ws["H1"].Style.DiagonalBorder.SetBorderStyle("#0000FF", BorderType.Thin);
            var rule = ws.DataValidations.AddIntegerRule("C1:C3", DvOp.Between, 1, 9);
            rule.ComparisonOperator = DvOp.NotBetween;
            ws.GetCellAt(0, 0).Hyperlink = "https://example.com/one";
            ws.GetCellAt(0, 0).Hyperlink = "mailto:me@example.com";
            wb.SaveAs(Tmp("edited.xlsx"));

            using (var doc = _excel.Open(Tmp("edited.xlsx")))
            {
                dynamic xs = doc.Sheet(1);
                // Excel's xlDiagonalDown is 5 and xlDiagonalUp is 6; OpenSheet's Forward is the upward one, Backward the downward one
                int Down(string a) => (int)xs.Range(a).Borders.Item(5).LineStyle;
                int Up(string a) => (int)xs.Range(a).Borders.Item(6).LineStyle;
                Assert.True(Down("D1") != None && Up("D1") != None, "Both draws both diagonals");
                Assert.True(Down("E1") == None && Up("E1") != None, "Forward draws the upward one only");
                Assert.True(Down("F1") == None && Up("F1") == None, "a cell that shares the border but not the direction draws none");
                Assert.True(Down("H1") != None && Up("H1") == None, "a direction set before the line exists is kept");

                Assert.Equal(2, (int)xs.Range("C1").Validation.Operator);          // xlNotBetween
                Assert.Equal(1, (int)xs.Hyperlinks.Count);
                Assert.Equal("mailto:me@example.com", (string)xs.Hyperlinks.Item(1).Address);
            }
        }

        [ExcelFact]
        public void ResizedPictures_HaveTheSizeExcelDrawsThem()
        {
            byte[] Picture(int w, int h)
            {
                using (var img = new SixLabors.ImageSharp.Image<SixLabors.ImageSharp.PixelFormats.Rgba32>(w, h))
                using (var ms = new System.IO.MemoryStream()) { SixLabors.ImageSharp.ImageExtensions.SaveAsPng(img, ms); return ms.ToArray(); }
            }
            var wb = WorkBook.Create(); var ws = wb.DefaultWorkSheet;
            ws.InsertImage(Picture(96, 48), ImageFormat.PNG, 1, 1, 2, 2).Resize();
            ws.InsertImage(Picture(96, 48), ImageFormat.PNG, 5, 5, 6, 6).Resize(2.5);
            ws.InsertImage(Picture(200, 100), ImageFormat.PNG, 9, 9, 10, 10).Resize(0.5, 0.25);
            wb.SaveAs(Tmp("pictures.xlsx"));

            using (var doc = _excel.Open(Tmp("pictures.xlsx")))
            {
                dynamic xs = doc.Sheet(1);
                // a pixel is 0.75 point
                var expected = new[] { (96 * 0.75, 48 * 0.75), (240 * 0.75, 120 * 0.75), (100 * 0.75, 25 * 0.75) };
                for (var i = 0; i < expected.Length; i++)
                {
                    dynamic shape = xs.Shapes.Item(i + 1);
                    Assert.InRange((double)shape.Width, expected[i].Item1 - 1, expected[i].Item1 + 1);
                    Assert.InRange((double)shape.Height, expected[i].Item2 - 1, expected[i].Item2 + 1);
                }
            }
        }

        string WithChartsAndPicture()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet; ws.Name = "Charts";
            for (var i = 0; i < 4; i++) { ws.GetCellAt(i, 0).Value = "k" + i; ws.GetCellAt(i, 1).Value = i + 1; }
            var a = ws.CreateChart(ChartType.Column, 6, 0, 16, 5); a.AddSeries("A1:A4", "B1:B4"); a.Plot();
            var b = ws.CreateChart(ChartType.Line, 18, 0, 28, 5); b.AddSeries("A1:A4", "B1:B4"); b.Plot();
            ws.InsertImage(Samples.Png, ImageFormat.PNG, 7, 0, 9, 3);
            wb.CreateWorkSheet("Plain")["A1"].Value = "stays";
            wb.SaveAs(Tmp("start.xlsx"));
            return Tmp("start.xlsx");
        }

        [ExcelFact]
        public void RemovedCharts_AndDeletedAndReplacedSheets_LeaveFilesExcelOpensWithTheRightContent()
        {
            var start = WithChartsAndPicture();

            var removed = WorkBook.Load(start);
            removed.DefaultWorkSheet.RemoveChart(removed.DefaultWorkSheet.Charts.Last());
            removed.SaveAs(Tmp("chart-removed.xlsx"));

            var deleted = WorkBook.Load(start);
            deleted.RemoveWorkSheet("Charts");
            deleted.GetWorkSheet("Plain").InsertImage(Samples.Png, ImageFormat.PNG, 1, 1, 3, 3);   // takes the place of the deleted sheet's drawing
            deleted.SaveAs(Tmp("deleted.xlsx"));

            var replaced = WorkBook.Load(start);
            var fresh = WorkBook.Create(); fresh.DefaultWorkSheet.Name = "Fresh"; fresh.DefaultWorkSheet["A1"].Value = "new";
            fresh.DefaultWorkSheet.InsertImage(Samples.Png, ImageFormat.PNG, 1, 1, 3, 3);
            replaced.WorkSheets[0] = fresh.DefaultWorkSheet;
            replaced.SaveAs(Tmp("replaced.xlsx"));

            using (var doc = _excel.Open(Tmp("chart-removed.xlsx")))
            {
                dynamic xs = doc.Sheet("Charts");
                Assert.Equal(1, (int)xs.ChartObjects().Count);
                Assert.Equal(2, (int)xs.Shapes.Count);                             // the remaining chart and the picture
            }
            using (var doc = _excel.Open(Tmp("deleted.xlsx")))
            {
                Assert.Equal(1, (int)doc.Book.Worksheets.Count);
                dynamic xs = doc.Sheet("Plain");
                Assert.Equal("stays", (string)xs.Range("A1").Value2);
                Assert.Equal(1, (int)xs.Shapes.Count);                             // only the new picture; nothing of the deleted sheet's charts
                Assert.Equal(0, (int)xs.ChartObjects().Count);
            }
            using (var doc = _excel.Open(Tmp("replaced.xlsx")))
            {
                Assert.Equal(new[] { "Fresh", "Plain" }, Enumerable.Range(1, (int)doc.Book.Worksheets.Count).Select(i => (string)doc.Sheet(i).Name).ToArray());
                dynamic xs = doc.Sheet("Fresh");
                Assert.Equal("new", (string)xs.Range("A1").Value2);
                Assert.Equal(1, (int)xs.Shapes.Count);
            }
        }
    }
}
