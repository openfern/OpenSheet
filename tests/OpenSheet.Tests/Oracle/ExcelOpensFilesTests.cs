using System;
using Xunit;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Real Excel opens what OpenSheet writes, without any repair prompt, and sees the features we put in.
    /// This closes the gap the schema validator cannot: a file can be schema-valid and still be refused by Excel.
    /// </summary>
    [Collection("Excel")]
    public class ExcelOpensFilesTests : TestBase
    {
        readonly ExcelOracle _excel;
        public ExcelOpensFilesTests(ExcelOracle excel) { _excel = excel; }

        [ExcelFact]
        public void KitchenSink_Xlsx_OpensWithoutRepair_AndShowsEveryFeature()
        {
            Samples.KitchenSink().SaveAs(Tmp("sink.xlsx"));
            using (var doc = _excel.Open(Tmp("sink.xlsx")))
            {
                var d = new Diffs();
                dynamic wb = doc.Book;
                dynamic ws = doc.Sheet("Data");

                d.Eq("sheet count", 2, (int)wb.Worksheets.Count);
                d.Eq("A1", "month", (string)ws.Range("A1").Value2);
                d.Eq("B3", 20.0, (double)ws.Range("B3").Value2);
                d.Eq("A1 bold", true, (bool)ws.Range("A1").Font.Bold);

                d.Try("merge", () =>
                {
                    d.Eq("E1 merged", true, (bool)ws.Range("E1").MergeCells);
                    d.Eq("E1 merge area", "$E$1:$G$1", (string)ws.Range("E1").MergeArea.Address);
                });
                d.Try("autofilter", () => d.Eq("AutoFilterMode", true, (bool)ws.AutoFilterMode));
                d.Try("hyperlink", () =>
                {
                    d.Eq("hyperlink count", 1, (int)ws.Hyperlinks.Count);
                    d.Eq("hyperlink address", "https://example.com", ((string)ws.Range("A1").Hyperlinks.Item(1).Address).TrimEnd('/'));
                });
                d.Try("comment", () =>
                {
                    d.Eq("comment count", 1, (int)ws.Comments.Count);
                    d.Eq("comment text", "check this", ((string)ws.Range("B2").Comment.Text()).Trim());
                    d.Eq("comment author", "me", (string)ws.Range("B2").Comment.Author);
                    d.Eq("comment hidden", false, (bool)ws.Range("B2").Comment.Visible);
                });
                d.Try("named range", () =>
                {
                    dynamic n = wb.Names.Item("Months");
                    d.Eq("name refers to", "=Data!$A$1:$A$6", (string)n.RefersTo);
                });
                // The file stores width/256 characters including 5px of padding; Excel's ColumnWidth leaves the padding out (about 0.71 chars at Calibri 11).
                d.Try("column width", () => d.True("col A width ~ 4000/256 chars", Math.Abs((double)ws.Columns(1).ColumnWidth + 5.0 / 7 - 4000 / 256.0) < 0.15));
                d.Try("tab color", () => d.Eq("tab color (BGR)", 255, (int)ws.Tab.Color));
                d.Try("outline", () =>
                {
                    d.Eq("row 3 outline", 2, (int)ws.Rows(3).OutlineLevel);
                    d.Eq("row 4 outline", 2, (int)ws.Rows(4).OutlineLevel);
                    d.Eq("row 5 outline", 1, (int)ws.Rows(5).OutlineLevel);
                });

                d.Try("table", () =>
                {
                    d.Eq("table count", 1, (int)ws.ListObjects.Count);
                    dynamic t = ws.ListObjects.Item(1);
                    d.Eq("table name", "Sales", (string)t.Name);
                    d.Eq("table range", "$R$1:$T$4", (string)t.Range.Address);
                    d.Eq("table style", "TableStyleMedium2", (string)t.TableStyle.Name);
                    d.Eq("table header", true, (bool)t.ShowHeaders);
                });

                d.Try("validation list", () =>
                {
                    dynamic v = ws.Range("H1").Validation;
                    d.Eq("H1 type (list)", 3, (int)v.Type);
                    d.Eq("H1 list", "a,b", ((string)v.Formula1).Replace(';', ',')); // Excel shows the Windows list separator
                });
                d.Try("validation integer", () =>
                {
                    dynamic v = ws.Range("I3").Validation;
                    d.Eq("I3 type (whole)", 1, (int)v.Type);
                    d.Eq("I3 operator (between)", 1, (int)v.Operator);
                    d.Eq("I3 min", "1", (string)v.Formula1);
                    d.Eq("I3 max", "10", (string)v.Formula2);
                    d.Eq("I3 error text", "1-10", (string)v.ErrorMessage);
                });
                d.Try("validation decimal", () => { dynamic v = ws.Range("J1").Validation; d.Eq("J1 type (decimal)", 2, (int)v.Type); d.Eq("J1 operator (>)", 5, (int)v.Operator); });
                d.Try("validation date", () => { dynamic v = ws.Range("K1").Validation; d.Eq("K1 type (date)", 4, (int)v.Type); d.Eq("K1 operator (>=)", 7, (int)v.Operator); });
                d.Try("validation text length", () => { dynamic v = ws.Range("L1").Validation; d.Eq("L1 type (length)", 6, (int)v.Type); d.Eq("L1 operator (<=)", 8, (int)v.Operator); });
                d.Try("validation formula list", () => { dynamic v = ws.Range("M1").Validation; d.Eq("M1 type (list)", 3, (int)v.Type); d.Eq("M1 source", "=$A$2:$A$6", (string)v.Formula1); });

                d.Try("conditional format", () =>
                {
                    dynamic fc = ws.Range("B2:B6").FormatConditions;
                    d.Eq("cf count", 1, (int)fc.Count);
                    dynamic r = fc.Item(1);
                    d.Eq("cf type (cell value)", 1, (int)r.Type);
                    d.Eq("cf operator (>)", 5, (int)r.Operator);
                    d.Eq("cf formula", "=20", (string)r.Formula1);
                    d.Eq("cf bold", true, (bool)r.Font.Bold);
                    d.Eq("cf font color (BGR red)", 255, (int)r.Font.Color);
                    d.Eq("cf fill (BGR yellow)", 65535, (int)r.Interior.Color);
                });

                d.Try("picture", () =>
                {
                    // the comment on B2 is also a shape, so look for pictures only
                    dynamic pic = null; var pictures = 0;
                    for (var i = 1; i <= (int)ws.Shapes.Count; i++)
                        if ((int)ws.Shapes.Item(i).Type == 13) { pictures++; pic = ws.Shapes.Item(i); }
                    d.Eq("picture count", 1, pictures);
                    if (pic != null) d.Eq("picture top-left", "$J$9", (string)pic.TopLeftCell.Address);
                });

                d.Try("page setup", () =>
                {
                    dynamic ps = ws.PageSetup;
                    d.Eq("orientation (landscape)", 2, (int)ps.Orientation);
                    d.Eq("paper (A4)", 9, (int)ps.PaperSize);
                    d.Eq("print area", "$A$1:$G$20", (string)ps.PrintArea);
                    d.Eq("title rows", "$1:$1", (string)ps.PrintTitleRows);
                    d.Eq("center header", "Report", (string)ps.CenterHeader);
                    d.Eq("right footer", "Page", (string)ps.RightFooter);
                });

                d.Try("properties", () =>
                {
                    d.Eq("author", "tester", (string)wb.BuiltinDocumentProperties.Item("Author").Value);
                    d.Eq("title", "Sink", (string)wb.BuiltinDocumentProperties.Item("Title").Value);
                    d.Eq("custom property", "v", (string)wb.CustomDocumentProperties.Item("k").Value);
                });

                d.Assert();
            }
        }

        [ExcelFact]
        public void KitchenSink_Charts_AreWhatWeAskedFor()
        {
            Samples.KitchenSink().SaveAs(Tmp("sink.xlsx"));
            using (var doc = _excel.Open(Tmp("sink.xlsx")))
            {
                var d = new Diffs();
                dynamic ws = doc.Sheet("Charts");
                d.Eq("chart count", 6, (int)ws.ChartObjects().Count);
                // xlColumnClustered 51, xlBarClustered 57, xlLine 4, xlArea 1, xlPie 5, xlXYScatter -4169
                var expected = new (string title, int type, int series)[] { ("Column", 51, 2), ("Bar", 57, 2), ("Line", 4, 2), ("Area", 1, 2), ("Pie", 5, 1) };
                for (var i = 0; i < expected.Length; i++)
                {
                    var n = i + 1; var e = expected[i];
                    d.Try("chart " + e.title, () =>
                    {
                        dynamic co = ws.ChartObjects(n);
                        d.Eq(e.title + " type", e.type, (int)co.Chart.ChartType);
                        d.Eq(e.title + " series count", e.series, (int)co.Chart.SeriesCollection().Count);
                        d.Eq(e.title + " has title", true, (bool)co.Chart.HasTitle);
                        d.Eq(e.title + " title", e.title, (string)co.Chart.ChartTitle.Text);
                        d.Eq(e.title + " has legend", true, (bool)co.Chart.HasLegend);
                        d.Eq(e.title + " legend (bottom)", -4107, (int)co.Chart.Legend.Position);
                        d.Eq(e.title + " first series name", "Sales", (string)co.Chart.SeriesCollection(1).Name);
                        d.Eq(e.title + " top-left", "$A$" + (i * 16 + 1), (string)co.TopLeftCell.Address);
                    });
                }
                d.Try("scatter", () =>
                {
                    dynamic co = ws.ChartObjects(6);
                    d.Eq("scatter type", -4169, (int)co.Chart.ChartType);
                    d.Eq("scatter series", 1, (int)co.Chart.SeriesCollection().Count);
                });
                d.Try("series formula", () =>
                {
                    var f = (string)ws.ChartObjects(1).Chart.SeriesCollection(1).Formula;
                    d.True("series 1 formula points at Data!B2:B6, got " + f, f.Contains("Data!$B$2:$B$6"));
                    d.True("series 1 categories are Data!A2:A6, got " + f, f.Contains("Data!$A$2:$A$6"));
                });
                d.Assert();
            }
        }

        [ExcelFact]
        public void Xls_Opens_AndKeepsValuesStylesAndFormulas()
        {
            var wb = WorkBook.Create(ExcelFileFormat.XLS);
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "xls"; ws["A2"].Value = 2.5; ws["A3"].Formula = "=A2*4"; ws["A4"].Value = new DateTime(2024, 3, 9);
            ws["A1"].Style.Font.Bold = true; ws["A2"].FormatString = "0.00";
            wb.SaveAs(Tmp("old.xls"));
            using (var doc = _excel.Open(Tmp("old.xls")))
            {
                var d = new Diffs();
                dynamic s = doc.Sheet(1);
                d.Eq("A1", "xls", (string)s.Range("A1").Value2);
                d.Eq("A2", 2.5, (double)s.Range("A2").Value2);
                d.Eq("A3 formula", "=A2*4", (string)s.Range("A3").Formula);
                d.Eq("A3 value", 10.0, (double)s.Range("A3").Value2);
                d.Eq("A4 serial", new DateTime(2024, 3, 9).ToOADate(), (double)s.Range("A4").Value2);
                d.Eq("A1 bold", true, (bool)s.Range("A1").Font.Bold);
                d.Eq("A2 format", "0.00", (string)s.Range("A2").NumberFormat);
                d.Assert();
            }
        }

        [ExcelFact]
        public void Encrypted_Opens_OnlyWithThePassword()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = "secret";
            wb.SaveAs(Tmp("locked.xlsx"), "pa55word");

            using (var doc = _excel.Open(Tmp("locked.xlsx"), "pa55word"))
                Assert.Equal("secret", (string)doc.Sheet(1).Range("A1").Value2);

            // Opening with no password at all would make Excel show a modal prompt and hang, so only a wrong one is tried.
            Assert.ThrowsAny<Exception>(() => _excel.Open(Tmp("locked.xlsx"), "wrong").Dispose());
        }

        [ExcelFact]
        public void Reloaded_AndResavedFile_StillOpens()
        {
            Samples.KitchenSink().SaveAs(Tmp("one.xlsx"));
            var again = WorkBook.Load(Tmp("one.xlsx"));
            again.DefaultWorkSheet["Z1"].Value = "edited";
            again.SaveAs(Tmp("two.xlsx"));
            using (var doc = _excel.Open(Tmp("two.xlsx")))
            {
                Assert.Equal("edited", (string)doc.Sheet("Data").Range("Z1").Value2);
                Assert.Equal(6, (int)doc.Sheet("Charts").ChartObjects().Count);
            }
        }
    }
}
