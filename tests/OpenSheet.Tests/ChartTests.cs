using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OpenSheet.Drawing.Charts;
using OpenSheet.Drawing.Images;
using Xunit;

namespace OpenSheet.Tests
{
    public class ChartTests : TestBase
    {
        static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        WorkSheet Data(WorkBook wb)
        {
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "month"; ws["B1"].Value = "sales"; ws["C1"].Value = "cost";
            var months = new[] { "Jan", "Feb", "Mar", "Apr" };
            for (var i = 0; i < 4; i++) { ws.GetCellAt(i + 1, 0).Value = months[i]; ws.GetCellAt(i + 1, 1).Value = 10 * (i + 1); ws.GetCellAt(i + 1, 2).Value = 7 * (i + 1); }
            return ws;
        }

        static XDocument Part(string xlsx, string entry)
        {
            using (var z = ZipFile.OpenRead(xlsx)) using (var s = z.GetEntry(entry).Open()) return XDocument.Load(s);
        }

        [Theory]
        [InlineData(ChartType.Column, "barChart", "col")]
        [InlineData(ChartType.Bar, "barChart", "bar")]
        [InlineData(ChartType.Line, "lineChart", null)]
        [InlineData(ChartType.Area, "areaChart", null)]
        [InlineData(ChartType.Pie, "pieChart", null)]
        public void EachType_WritesSeriesWithReferencesAndCaches(ChartType type, string element, string barDir)
        {
            var wb = WorkBook.Create();
            var ws = Data(wb);
            var chart = ws.CreateChart(type, 7, 0, 20, 6);
            var s1 = chart.AddSeries("A2:A5", "B2:B5"); s1.Title = "Sales";
            if (type != ChartType.Pie) chart.AddSeries("A2:A5", "C2:C5");
            chart.SetTitle("Report");
            chart.SetLegendPosition(LegendPosition.Bottom);
            chart.Plot();
            wb.SaveAs(Tmp("c.xlsx"));

            var doc = Part(Tmp("c.xlsx"), "xl/charts/chart1.xml");
            var plot = doc.Descendants(C + element).Single();
            var sers = plot.Elements(C + "ser").ToList();
            Assert.Equal(type == ChartType.Pie ? 1 : 2, sers.Count);
            Assert.Equal("Sheet1!$B$2:$B$5".Replace("Sheet1", "'Sheet1'"), sers[0].Element(C + "val").Descendants(C + "f").Single().Value);
            Assert.Equal("'Sheet1'!$A$2:$A$5", sers[0].Element(C + "cat").Descendants(C + "f").Single().Value);
            Assert.Equal("Sales", sers[0].Element(C + "tx").Element(C + "v").Value);
            Assert.Equal(new[] { "10", "20", "30", "40" }, sers[0].Element(C + "val").Descendants(C + "v").Select(v => v.Value).ToArray());
            Assert.Equal(new[] { "Jan", "Feb", "Mar", "Apr" }, sers[0].Element(C + "cat").Descendants(C + "v").Select(v => v.Value).ToArray());
            if (barDir != null) Assert.Equal(barDir, plot.Element(C + "barDir").Attribute("val").Value);
            Assert.Equal("b", doc.Descendants(C + "legendPos").Single().Attribute("val").Value);
            Assert.Contains("Report", doc.ToString());
            Assert.Equal(type == ChartType.Pie ? 0 : 1, doc.Descendants(C + "catAx").Count());

            var back = WorkBook.Load(Tmp("c.xlsx")).DefaultWorkSheet;
            Assert.Single(back.Charts);
        }

        [Fact]
        public void Scatter_UsesNumericXValues_AndTwoValueAxes()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            for (var i = 0; i < 5; i++) { ws.GetCellAt(i, 0).Value = i * 1.5; ws.GetCellAt(i, 1).Value = i * i; }
            var chart = ws.CreateChart(ChartType.Scatter, 6, 0, 16, 6);
            chart.AddSeries("A1:A5", "B1:B5");
            chart.Plot();
            wb.SaveAs(Tmp("s.xlsx"));
            var doc = Part(Tmp("s.xlsx"), "xl/charts/chart1.xml");
            var ser = doc.Descendants(C + "scatterChart").Single().Element(C + "ser");
            Assert.Equal(new[] { "0", "1.5", "3", "4.5", "6" }, ser.Element(C + "xVal").Descendants(C + "v").Select(v => v.Value).ToArray());
            Assert.Equal(2, doc.Descendants(C + "valAx").Count());
        }

        [Fact]
        public void SingleRangeSeries_GetsNumberedCategories()
        {
            var wb = WorkBook.Create();
            var ws = Data(wb);
            var chart = ws.CreateChart(ChartType.Column, 7, 0, 15, 5);
            chart.AddSeries("B2:B5");
            wb.SaveAs(Tmp("n.xlsx"));                        // never plotted: saving plots it
            var cat = Part(Tmp("n.xlsx"), "xl/charts/chart1.xml").Descendants(C + "cat").Single();
            Assert.Equal(new[] { "1", "2", "3", "4" }, cat.Descendants(C + "v").Select(v => v.Value).ToArray());
        }

        [Fact]
        public void CrossSheetSeries_AndErrors()
        {
            var wb = WorkBook.Create();
            var ws = Data(wb);
            var other = wb.CreateWorkSheet("Data 2");
            other["A1"].Value = 5; other["A2"].Value = 6;
            var chart = ws.CreateChart(ChartType.Line, 7, 0, 15, 5);
            chart.AddSeries("'Data 2'!A1:A2");
            chart.Plot();
            wb.SaveAs(Tmp("x.xlsx"));
            Assert.Equal("'Data 2'!$A$1:$A$2", Part(Tmp("x.xlsx"), "xl/charts/chart1.xml").Descendants(C + "val").Single().Descendants(C + "f").Single().Value);

            Assert.Throws<ArgumentException>(() => chart.AddSeries("A1:B5"));                    // not one row/column
            Assert.Throws<ArgumentException>(() => chart.AddSeries("A1:A3", "B1:B4"));           // sizes differ
            Assert.Throws<ArgumentException>(() => chart.AddSeries("Nope!A1:A3"));
            Assert.Throws<InvalidOperationException>(() => ws.CreateChart(ChartType.Pie, 0, 0, 5, 5).Plot());
            Assert.Throws<NotSupportedException>(() => WorkBook.Create(ExcelFileFormat.XLS).DefaultWorkSheet.CreateChart(ChartType.Pie, 0, 0, 5, 5));
        }

        [Fact]
        public void LoadedChart_CanRetitle_ButNotChangeData()
        {
            var wb = WorkBook.Create();
            var ws = Data(wb);
            var c = ws.CreateChart(ChartType.Column, 7, 0, 15, 5);
            c.AddSeries("A2:A5", "B2:B5"); c.Plot();
            wb.SaveAs(Tmp("l.xlsx"));
            var back = WorkBook.Load(Tmp("l.xlsx"));
            var loaded = back.DefaultWorkSheet.Charts.Single();
            loaded.SetTitle("New title");
            loaded.SetLegendPosition(LegendPosition.Top);
            Assert.Throws<NotSupportedException>(() => loaded.AddSeries("B2:B5"));
            back.SaveAs(Tmp("l2.xlsx"));
            Assert.Contains("New title", Part(Tmp("l2.xlsx"), "xl/charts/chart1.xml").ToString());
        }

        [Fact]
        public void LoadedChart_ExposesItsPosition_AndEditsReachTheFile()
        {
            var wb = WorkBook.Create();
            var ws = Data(wb);
            var c = ws.CreateChart(ChartType.Column, 7, 1, 15, 6);
            c.AddSeries("A2:A5", "B2:B5"); c.Plot();
            wb.SaveAs(Tmp("p1.xlsx"));

            var back = WorkBook.Load(Tmp("p1.xlsx"));
            var pos = back.DefaultWorkSheet.Charts.Single().Position;
            Assert.Equal(1, pos.LeftColumnIndex); Assert.Equal(7, pos.TopRowIndex); Assert.Equal(6, pos.RightColumnIndex); Assert.Equal(15, pos.BottomRowIndex);
            pos.LeftColumnIndex = 3; pos.BottomRowIndex = 20;
            back.SaveAs(Tmp("p2.xlsx"));

            var again = WorkBook.Load(Tmp("p2.xlsx")).DefaultWorkSheet.Charts.Single().Position;
            Assert.Equal(3, again.LeftColumnIndex); Assert.Equal(20, again.BottomRowIndex);
        }

        [Fact]
        public void RemoveChart_RemovesItFromTheSheet()
        {
            var wb = WorkBook.Create();
            var ws = Data(wb);
            var a = ws.CreateChart(ChartType.Column, 7, 0, 15, 5); a.AddSeries("B2:B5"); a.Plot();
            var b = ws.CreateChart(ChartType.Line, 17, 0, 25, 5); b.AddSeries("B2:B5"); b.Plot();
            ws.RemoveChart(a);
            wb.SaveAs(Tmp("r.xlsx"));
            var back = WorkBook.Load(Tmp("r.xlsx")).DefaultWorkSheet;
            Assert.Single(back.Charts);
            Assert.Single(Part(Tmp("r.xlsx"), "xl/drawings/drawing1.xml").Descendants(Xdr + "graphicFrame"));
        }

        [Fact]
        public void Position_EditsReachTheFile_ForChartsAndPictures()
        {
            var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
            var wb = WorkBook.Create();
            var ws = Data(wb);
            var chart = ws.CreateChart(ChartType.Column, 7, 0, 15, 5);
            chart.AddSeries("B2:B5"); chart.Plot();
            chart.Position.RightColumnIndex = 11;
            chart.Position.BottomRowIndex = 30;
            var img = ws.InsertImage(png, ImageFormat.PNG, 1, 1, 3, 3);
            img.Position.RightColumnIndex = 9;
            wb.SaveAs(Tmp("p.xlsx"));

            var anchors = Part(Tmp("p.xlsx"), "xl/drawings/drawing1.xml").Descendants(Xdr + "twoCellAnchor").ToList();
            var frame = anchors.Single(a => a.Element(Xdr + "graphicFrame") != null);
            Assert.Equal("11", frame.Element(Xdr + "to").Element(Xdr + "col").Value);
            Assert.Equal("30", frame.Element(Xdr + "to").Element(Xdr + "row").Value);
            var pic = anchors.Single(a => a.Element(Xdr + "pic") != null);
            Assert.Equal("9", pic.Element(Xdr + "to").Element(Xdr + "col").Value);
        }
    }
}
