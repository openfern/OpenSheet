using System;
using System.Linq;
using OpenSheet.Options;
using Xunit;

namespace OpenSheet.Tests
{
    public class SheetSettingsTests : TestBase
    {
        [Theory]
        [InlineData("s.xlsx")]
        [InlineData("s.xls")]
        public void Sizes_Visibility_Grouping_Protection_RoundTrip(string file)
        {
            var wb = WorkBook.Create(file.EndsWith("xls") ? ExcelFileFormat.XLS : ExcelFileFormat.XLSX);
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "x"; ws["B2"].Value = "y";
            ws.GetColumn(1).Width = 6000; ws.GetColumn(2).Hidden = true;
            ws.GetRow(1).Height = 600;   ws.GetRow(2).Hidden = true;
            ws.GroupRows(3, 5); ws.GroupColumns("E", "G");
            ws.DisplayGridlines = false; ws.IsPrintGridlines = true; ws.FitToPage = true;
            ws.HorizontallyCenter = true; ws.VerticallyCenter = true; ws.IsRightToLeft = true;
            ws.SetZoom(150);
            ws.SetWorksheetTabColorIndex(10);
            ws.ProtectSheet("pw");
            var hidden = wb.CreateWorkSheet("Hide"); hidden.Hidden = HiddenState.VeryHidden;
            wb.SetActiveTab(0);
            wb.SaveAs(Tmp(file));

            var b = WorkBook.Load(Tmp(file));
            var s = b.DefaultWorkSheet;
            Assert.Equal(6000, s.GetColumn(1).Width);
            Assert.True(s.GetColumn(2).Hidden);
            Assert.False(s.GetColumn(0).Hidden);
            Assert.Equal(600, s.GetRow(1).Height);
            Assert.True(s.GetRow(2).Hidden);
            Assert.False(s.DisplayGridlines);
            Assert.True(s.IsPrintGridlines);
            Assert.True(s.FitToPage);
            Assert.True(s.HorizontallyCenter && s.VerticallyCenter);
            Assert.True(s.IsRightToLeft);
            Assert.Equal(HiddenState.VeryHidden, b.GetWorkSheet("Hide").Hidden);
            Assert.Equal(HiddenState.Visible, s.Hidden);
            Assert.Equal(0, s.Index);
        }

        [Fact]
        public void TabColor_Rgb_RoundTrips_InXlsx()
        {
            var wb = WorkBook.Create();
            Assert.Null(wb.DefaultWorkSheet.TabColor);
            wb.DefaultWorkSheet.SetWorksheetTabColor(new byte[] { 10, 20, 30 });
            wb.SaveAs(Tmp("t.xlsx"));
            Assert.Equal(new byte[] { 10, 20, 30 }, WorkBook.Load(Tmp("t.xlsx")).DefaultWorkSheet.TabColor);
            Assert.Throws<ArgumentException>(() => wb.DefaultWorkSheet.SetWorksheetTabColor(new byte[] { 1, 2 }));
            Assert.Throws<NotSupportedException>(() => WorkBook.Create(ExcelFileFormat.XLS).DefaultWorkSheet.SetWorksheetTabColor(new byte[] { 1, 2, 3 }));
        }

        [Fact]
        public void Protection_CanBeLifted()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws.ProtectSheet("pw");
            Assert.True(ws.Native.Protect);
            ws.UnprotectSheet();
            Assert.False(ws.Native.Protect);
        }

        [Fact]
        public void Ungroup_RowsAndColumns_DoNotThrow()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = 1;
            ws.GroupRows(1, 4); ws.UngroupRows(1, 4);
            ws.GroupColumns(1, 4); ws.UngroupColumns(1, 4);
            ws.GroupColumns("B", "D"); ws.UngroupColumn("B", "D");
        }

        [Fact]
        public void AutoSize_Widens_ForLongText()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = "a considerably longer piece of text than the default width";
            var before = ws.GetColumn(0).Width;
            ws.AutoSizeColumn(0);
            Assert.True(ws.GetColumn(0).Width > before);
            ws["B1"].Value = "multi\nline\ntext"; ws["B1"].Style.WrapText = true;
            ws.AutoSizeRow(0);
            ws.GetRow(0).AutoSizeRow(true);
            ws.GetColumn(1).AutoSizeColumn(true);
        }

        [Fact]
        public void Breaks_PrintTitles_AndAutoFilter()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            for (var r = 0; r < 5; r++) { ws.GetCellAt(r, 0).Value = r; ws.GetCellAt(r, 1).Value = "v" + r; }
            ws.SetRowBreak(2); ws.SetColumnBreak(1);
            ws.RemoveRowBreak(2); ws.RemoveColumnBreak(1);
            ws.SetRepeatingRows(0, 0); ws.SetRepeatingColumns("A", "A");
            ws.RemoveRepeatingRows(); ws.RemoveRepeatingColumns();
            ws.SetAutoFilter("A1:B5");
            wb.SaveAs(Tmp("f.xlsx"));
            ws.UnsetAutoFilter();
            wb.SaveAs(Tmp("g.xlsx"));
            using (var z = System.IO.Compression.ZipFile.OpenRead(Tmp("f.xlsx"))) using (var r = new System.IO.StreamReader(z.GetEntry("xl/worksheets/sheet1.xml").Open())) Assert.Contains("<autoFilter", r.ReadToEnd());
            using (var z = System.IO.Compression.ZipFile.OpenRead(Tmp("g.xlsx"))) using (var r = new System.IO.StreamReader(z.GetEntry("xl/worksheets/sheet1.xml").Open())) Assert.DoesNotContain("<autoFilter", r.ReadToEnd());
        }

        [Fact]
        public void SplitPane_And_RemovePane()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws.CreateSplitPane(2000, 1000, 1, 1, PanePosition.LowerRight);
            ws.RemovePane();
            ws.CreateFreezePane(2, 2, 3, 3);
        }

        [Fact]
        public void NamedTable_RoundTrip_WithAndWithoutHeader()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "id"; ws["B1"].Value = "name";
            ws["A2"].Value = 1; ws["B2"].Value = "x"; ws["A3"].Value = 2; ws["B3"].Value = "y";
            ws["A1:B3"].SaveAsNamedTable("People", true, Styles.TableStyle.TableStyleMedium9, true);
            ws["D1:E2"].SaveAsNamedTable("Raw");
            wb.SaveAs(Tmp("tb.xlsx"));

            var b = WorkBook.Load(Tmp("tb.xlsx")).DefaultWorkSheet;
            Assert.Equal(new[] { "People", "Raw" }, b.GetNamedTableNames().OrderBy(n => n).ToArray());
            Assert.Equal("A1:B3", b.GetNamedTable("People").RangeAddressAsString);
            var dt = b.GetDataTableFromNamedTable("People", true);
            Assert.Equal(new[] { "id", "name" }, dt.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToArray());
            Assert.Equal(2, dt.Rows.Count);
            Assert.Throws<ArgumentException>(() => b.GetNamedTable("Nope"));
            Assert.Throws<NotSupportedException>(() => WorkBook.Create(ExcelFileFormat.XLS).DefaultWorkSheet.GetNamedTableNames());
        }

        [Fact]
        public void Collections_Operations()
        {
            var wb = WorkBook.Create();
            var other = WorkBook.Create();
            other.DefaultWorkSheet["A1"].Value = "from other";
            other.DefaultWorkSheet.Name = "Imported";
            wb.WorkSheets.Add(other.DefaultWorkSheet);                 // copies into this workbook
            Assert.Equal("from other", wb.GetWorkSheet("Imported")["A1"].Value);
            Assert.Throws<ArgumentException>(() => wb.WorkSheets.Add(wb.DefaultWorkSheet));
            var inBook = wb.WorkSheets.Contains(wb.GetWorkSheet("Imported")); Assert.True(inBook);
            var foreign = wb.WorkSheets.Contains(other.DefaultWorkSheet); Assert.False(foreign);
            Assert.Equal(1, wb.WorkSheets.IndexOf(wb.GetWorkSheet("Imported")));
            Assert.Equal(1, wb.WorkSheets.FindIndex(s => s.Name == "Imported"));
            wb.WorkSheets.SetSheetIndex(wb.GetWorkSheet("Imported"), 0);
            Assert.Equal("Imported", wb.WorkSheets[0].Name);
            Assert.True(wb.WorkSheets.Remove(wb.GetWorkSheet("Imported")));
            Assert.Single(wb.WorkSheets);
            var arr = new WorkSheet[1]; wb.WorkSheets.CopyTo(arr, 0);
            Assert.Same(wb.DefaultWorkSheet, arr[0]);
            Assert.Throws<ArgumentOutOfRangeException>(() => wb.WorkSheets[9]);
        }

        [Fact]
        public void ToStream_And_Bytes_RoundTrip()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = "streamed";
            using (var ms = wb.ToStream()) Assert.Equal("streamed", WorkBook.Load(ms).DefaultWorkSheet["A1"].Value);
            Assert.Equal("streamed", WorkBook.Load(wb.ToByteArray()).DefaultWorkSheet["A1"].Value);
            Assert.Equal("streamed", WorkBook.Load(wb.ToXlsStream()).DefaultWorkSheet["A1"].Value);
            Assert.Equal("streamed", WorkBook.Load(wb.ToXlsxStream()).DefaultWorkSheet["A1"].Value);
            wb.Password = "pw";
            Assert.Equal("streamed", WorkBook.Load(wb.ToStream(), "pw").DefaultWorkSheet["A1"].Value);
            Assert.Equal("streamed", WorkBook.Load(wb.Encrypt("other"), "other").DefaultWorkSheet["A1"].Value);
            Assert.Equal("streamed", WorkBook.Load(WorkBook.Decrypt(wb.Encrypt(), "pw")).DefaultWorkSheet["A1"].Value);
        }

        [Fact]
        public void TextStreams_AndSheetExports()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "a"; ws["B1"].Value = 2; wb.CreateWorkSheet("Two")["A1"].Value = "z";
            using (var r = new System.IO.StreamReader(wb.ToCsvStream())) Assert.Equal("a,2\r\n", r.ReadToEnd());
            using (var r = new System.IO.StreamReader(wb.ToTsvStream(1))) Assert.Equal("z\r\n", r.ReadToEnd());
            Assert.Equal(2, wb.ToCsvStreams(";").Count());
            Assert.Equal(2, wb.ToJsonStreams().Count());
            Assert.Equal(2, wb.ToXmlStreams().Count());
            Assert.Equal(2, wb.ToTsvStreams().Count());
            ws.SaveAs(Tmp("sheet.csv"));
            ws.SaveAsNewExcelFile(Tmp("sheet.xls"), ExcelFileFormat.XLS);
            Assert.Equal("a", WorkBook.Load(Tmp("sheet.xls")).DefaultWorkSheet["A1"].Value);
            Assert.Equal("a,2", System.IO.File.ReadAllText(Tmp("sheet.csv")).Trim());
            using (var ms = ws.ToNewExcelStream()) Assert.Equal("a", WorkBook.Load(ms).DefaultWorkSheet["A1"].Value);
            wb.SaveAsCsv(Tmp("named.txt"), ";");                             // explicit method decides the format, not the extension
            Assert.Equal("a;2", System.IO.File.ReadAllText(Tmp("named.txt")).Trim());
            Assert.True(System.IO.File.Exists(Tmp("named.Two.txt")));         // second sheet gets its own file
        }

        [Fact]
        public void ClearContents_KeepsFormatting_ButValueNullRemovesCells()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = "x"; ws["A1"].Style.Font.Bold = true;
            ws["A1"].ClearContents();
            Assert.True(ws.GetCellAt(0, 0).IsEmpty);
            Assert.True(ws["A1"].Style.Font.Bold);          // style survived
            ws["A1"].Value = null;
            Assert.False(ws["A1"].Style.Font.Bold);         // cell gone, default style again
        }

        [Fact]
        public void FilledCells_And_UsedRange()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            Assert.Null(ws.FirstFilledCell);
            Assert.Equal("A1", ws.RangeAddressAsString);
            ws["C3"].Value = 1; ws["E2"].Value = 2;
            Assert.Equal("C2:E3", ws.RangeAddressAsString);
            Assert.Equal("E2", ws.FirstFilledCell.AddressString);
            Assert.Equal("C3", ws.LastFilledCell.AddressString);
            Assert.Equal(2, ws.FilledCells.Count);
            Assert.Equal(2, ws.PhysicalRowCount);
            Assert.Equal(2, ws.PhysicalColumnCount);
            Assert.Equal(3, ws.ToMultiDimensionalArray()[0].Length);
        }
    }
}
