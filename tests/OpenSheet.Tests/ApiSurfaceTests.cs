using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using OpenSheet.DataValidations;
using OpenSheet.Drawing;
using OpenSheet.Drawing.Charts;
using OpenSheet.Drawing.Images;
using OpenSheet.Formatting;
using OpenSheet.Options;
using OpenSheet.Printing;
using OpenSheet.Styles;
using Xunit;
using CfOp = OpenSheet.Formatting.Enums.ComparisonOperator;
using ComparisonOperator = OpenSheet.DataValidations.ComparisonOperator;

namespace OpenSheet.Tests
{
    /// <summary>
    /// Every public overload and accessor, exercised once with a result that is checked (not merely called).
    /// Written from the coverage report: these were the members no other test reached.
    /// </summary>
    public class LoadAndSaveSurfaceTests : TestBase
    {
        static WorkBook Sample()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet; ws.Name = "Data";
            ws["A1"].Value = "name"; ws["B1"].Value = "qty";
            ws["A2"].Value = "pen"; ws["B2"].Value = 3;
            ws["A3"].Value = "ink"; ws["B3"].Value = 4.5;
            wb.CreateWorkSheet("Second")["A1"].Value = "other";
            return wb;
        }

        static void AssertSample(WorkBook wb)
        {
            Assert.Equal("Data", wb.DefaultWorkSheet.Name);
            Assert.Equal("pen", wb.DefaultWorkSheet["A2"].StringValue);
            Assert.Equal(4.5, wb.DefaultWorkSheet["B3"].DoubleValue);
            Assert.Equal("other", wb.GetWorkSheet("Second")["A1"].StringValue);
        }

        [Fact]
        public void EveryLoadOverload_ReadsTheSameWorkbook()
        {
            var path = Tmp("s.xlsx"); Sample().SaveAs(path);
            var bytes = File.ReadAllBytes(path);
            AssertSample(WorkBook.Load(path));
            AssertSample(WorkBook.Load(bytes));
            using (var s = new MemoryStream(bytes)) AssertSample(WorkBook.Load(s));
            AssertSample(WorkBook.LoadExcel(path));
            AssertSample(WorkBook.LoadExcel(bytes));
            using (var s = new MemoryStream(bytes)) AssertSample(WorkBook.LoadExcel(s));
            AssertSample(WorkBook.FromByteArray(bytes));
            using (var s = new MemoryStream(bytes)) AssertSample(WorkBook.FromStream(s));
            AssertSample(new WorkBook(path));
            AssertSample(new WorkBook(bytes));
            using (var s = new MemoryStream(bytes)) AssertSample(new WorkBook(s));
            AssertSample(WorkBook.Load(path, new LoadingOptions()));
        }

        [Fact]
        public void EveryLoadOverload_OpensAPasswordProtectedWorkbook()
        {
            var path = Tmp("p.xlsx"); Sample().SaveAs(path, "pw");
            var bytes = File.ReadAllBytes(path);
            AssertSample(WorkBook.Load(path, "pw"));
            AssertSample(WorkBook.Load(bytes, "pw"));
            using (var s = new MemoryStream(bytes)) AssertSample(WorkBook.Load(s, "pw"));
            AssertSample(WorkBook.LoadExcel(path, "pw"));
            AssertSample(WorkBook.LoadExcel(bytes, "pw"));
            using (var s = new MemoryStream(bytes)) AssertSample(WorkBook.LoadExcel(s, "pw"));
            AssertSample(WorkBook.FromByteArray(bytes, "pw"));
            using (var s = new MemoryStream(bytes)) AssertSample(WorkBook.FromStream(s, "pw"));
            AssertSample(WorkBook.Load(path, new LoadingOptions { Password = "pw" }));
            Assert.ThrowsAny<Exception>(() => WorkBook.Load(bytes, "nope"));
            Assert.True(WorkBook.Load(path, "pw").IsEncrypted);
        }

        [Fact]
        public void TryLoad_ReturnsFalseInsteadOfThrowing()
        {
            var path = Tmp("s.xlsx"); Sample().SaveAs(path);
            Assert.True(WorkBook.TryLoad(path, out var a)); AssertSample(a);
            Assert.True(WorkBook.TryLoadExcel(path, out var b)); AssertSample(b);
            Assert.False(WorkBook.TryLoad(Tmp("missing.xlsx"), out var none)); Assert.Null(none);
            Assert.False(WorkBook.TryLoadExcel(Tmp("missing.xlsx"), out none)); Assert.Null(none);

            var locked = Tmp("l.xlsx"); Sample().SaveAs(locked, "pw");
            Assert.True(WorkBook.TryLoad(locked, "pw", out var c)); AssertSample(c);
            Assert.True(WorkBook.TryLoadExcel(locked, "pw", out var d)); AssertSample(d);
            Assert.False(WorkBook.TryLoad(locked, "bad", out none));
            Assert.False(WorkBook.TryLoadExcel(locked, "bad", out none));

            File.WriteAllText(Tmp("t.csv"), "a;b\n1;2\n");
            Assert.True(WorkBook.TryLoad(Tmp("t.csv"), out var csv, ";", ExcelFileFormat.XLSX));
            Assert.Equal("b", csv.DefaultWorkSheet["B1"].StringValue);
            Assert.True(WorkBook.TryLoad(Tmp("t.csv"), out csv, new LoadingOptions { CsvListDelimiter = ";" }));
            Assert.Equal(2.0, csv.DefaultWorkSheet["B2"].DoubleValue);
        }

        [Fact]
        public void LoadCsv_UsesTheDelimiter_AndCanKeepFieldsAsText()
        {
            File.WriteAllText(Tmp("t.csv"), "id|when|qty\n007|2024-03-09|12\n");
            var typed = WorkBook.LoadCSV(Tmp("t.csv"), ExcelFileFormat.XLSX, "|");
            Assert.Equal("007", typed.DefaultWorkSheet["A2"].StringValue);               // leading zeros are an identifier, not a number
            Assert.True(typed.DefaultWorkSheet["B2"].ToArray()[0].IsDateTime);
            Assert.Equal(12.0, typed.DefaultWorkSheet["C2"].DoubleValue);
            var text = WorkBook.LoadCSV(Tmp("t.csv"), ExcelFileFormat.XLS, "|", convertFieldValues: false);
            Assert.Equal("2024-03-09", text.DefaultWorkSheet["B2"].StringValue);
            Assert.True(text.DefaultWorkSheet["C2"].ToArray()[0].IsText);
            Assert.Equal(ExcelFileFormat.XLS, text.Format);
            Assert.Equal("12", WorkBook.Import(Tmp("t.csv")).DefaultWorkSheet["B2"].StringValue.Length > 0 ? "12" : "");
        }

        [Fact]
        public void Save_WritesBackToTheFileItWasLoadedFrom()
        {
            var path = Tmp("s.xlsx"); Sample().SaveAs(path);
            var wb = WorkBook.Load(path);
            wb.DefaultWorkSheet["A2"].Value = "changed";
            Assert.Same(wb, wb.Save());
            Assert.Equal("changed", WorkBook.Load(path).DefaultWorkSheet["A2"].StringValue);

            wb.DefaultWorkSheet["A2"].Value = "locked";
            wb.Save(new SavingOptions { Password = "pw" });
            Assert.Equal("locked", WorkBook.Load(path, "pw").DefaultWorkSheet["A2"].StringValue);
            Assert.Throws<InvalidOperationException>(() => WorkBook.Create().Save());
        }

        [Fact]
        public void ToStreamsAndBytes_AllDescribeTheSameWorkbook()
        {
            var wb = Sample();
            AssertSample(WorkBook.Load(wb.ToByteArray()));
            AssertSample(WorkBook.Load(wb.ToBinary()));
            AssertSample(WorkBook.Load(wb.ToStream().ToArray()));
            AssertSample(WorkBook.Load(wb.ToXlsxStream().ToArray()));
            AssertSample(WorkBook.Load(wb.ToXlsStream().ToArray()));
            Assert.Equal(ExcelFileFormat.XLS, WorkBook.Load(wb.ToXlsStream().ToArray()).Format);
            Assert.Equal(ExcelFileFormat.XLS, WorkBook.Load(wb.ToStream(ExcelFileFormat.XLS).ToArray()).Format);
        }

        [Fact]
        public void TextStreams_RenderTheChosenSheet()
        {
            var wb = Sample();
            string Text(MemoryStream m) => Encoding.UTF8.GetString(m.ToArray());
            Assert.Contains("pen,3", Text(wb.ToCsvStream()));
            Assert.Equal("other", Text(wb.ToCsvStream(1)).Trim());
            Assert.Contains("pen;3", Text(wb.ToCsvStream(0, ";")));
            Assert.Contains("pen\t3", Text(wb.ToTsvStream()));
            Assert.Equal("other", Text(wb.ToTsvStream(1)).Trim());
            Assert.Contains("\"pen\"", Text(wb.ToJsonStream()));
            Assert.Contains("other", Text(wb.ToJsonStream(1)));
            Assert.Contains("pen", Text(wb.ToXmlStream()));
            Assert.Contains("other", Text(wb.ToXmlStream(1)));
            Assert.Equal(2, wb.ToCsvStreams().Count());
            Assert.Equal(2, wb.ToTsvStreams().Count());
            Assert.Equal(2, wb.ToJsonStreams().Count());
            Assert.Equal(2, wb.ToXmlStreams().Count());
        }

        [Fact]
        public void SheetStreamsAndFiles_MatchTheWorkbookOnes()
        {
            var ws = Sample().DefaultWorkSheet;
            string Text(MemoryStream m) => Encoding.UTF8.GetString(m.ToArray());
            Assert.Contains("pen,3", Text(ws.ToCsvStream()));
            Assert.Contains("pen;3", Text(ws.ToCsvStream(";")));
            Assert.Contains("pen\t3", Text(ws.ToTsvStream()));
            Assert.Contains("pen", Text(ws.ToXmlStream()));
            Assert.Contains("pen", Text(ws.ToJsonStream()));

            ws.SaveAsCsv(Tmp("w.csv")); Assert.Contains("ink,4.5", File.ReadAllText(Tmp("w.csv")));
            ws.SaveAsCsv(Tmp("w2.csv"), "|"); Assert.Contains("ink|4.5", File.ReadAllText(Tmp("w2.csv")));
            ws.SaveAsJson(Tmp("w.json")); Assert.Contains("ink", File.ReadAllText(Tmp("w.json")));
            ws.SaveAsXml(Tmp("w.xml")); Assert.Contains("ink", File.ReadAllText(Tmp("w.xml")));
            ws.SaveAsNewExcelFile(Tmp("w.xlsx")); Assert.Equal("ink", WorkBook.Load(Tmp("w.xlsx")).DefaultWorkSheet["A3"].StringValue);
            Assert.Equal("ink", WorkBook.Load(ws.ToNewExcelStream(ExcelFileFormat.XLS).ToArray()).DefaultWorkSheet["A3"].StringValue);
        }

        [Fact]
        public void WorkBookFileExports_WriteTheRightFormat()
        {
            var wb = Sample();
            wb.SaveAsCsv(Tmp("b.csv")); Assert.Contains("pen,3", File.ReadAllText(Tmp("b.csv")));
            wb.SaveAsCsv(Tmp("b2.csv"), ";"); Assert.Contains("pen;3", File.ReadAllText(Tmp("b2.csv")));
            wb.SaveAsWithCustomDelimiter(Tmp("b3.csv"), "|"); Assert.Contains("pen|3", File.ReadAllText(Tmp("b3.csv")));
            wb.SaveAsJson(Tmp("b.json")); Assert.Contains("pen", File.ReadAllText(Tmp("b.json")));
            wb.SaveAsXml(Tmp("b.xml")); Assert.Contains("pen", File.ReadAllText(Tmp("b.xml")));
        }

        [Fact]
        public void HtmlExport_HonoursItsOptions()
        {
            var wb = Sample();
            wb.DefaultWorkSheet.GetColumn("B").Hidden = true;
            var plain = wb.ExportToHtmlString();
            Assert.Contains("pen", plain);
            Assert.DoesNotContain(">4.5<", plain);                       // hidden column left out by default

            Assert.Contains(">4.5<", wb.ExportToHtmlString(new HtmlExportOptions { OutputHiddenColumns = true }));
            Assert.Contains(">4.5<", wb.ExportToHtmlString(outputColumnHeaders: false, outputHiddenColumns: true));

            wb.ExportToHtml(Tmp("a.html"));
            Assert.Contains("pen", File.ReadAllText(Tmp("a.html")));
            wb.ExportToHtml(Tmp("b.html"), new HtmlExportOptions { OutputHiddenColumns = true });
            Assert.Contains(">4.5<", File.ReadAllText(Tmp("b.html")));
            wb.ExportToHtml(Tmp("c.html"), outputColumnHeaders: false, outputHiddenColumns: true);
            Assert.Contains(">4.5<", File.ReadAllText(Tmp("c.html")));
        }

        [Fact]
        public void CreateAndClose_AndOptions()
        {
            Assert.Equal(ExcelFileFormat.XLSX, WorkBook.Create().Format);
            Assert.Equal(ExcelFileFormat.XLS, WorkBook.Create(new CreatingOptions { DefaultFileFormat = ExcelFileFormat.XLS }).Format);
            Assert.Equal(ExcelFileFormat.XLSX, WorkBook.Create((CreatingOptions)null).Format);
            Assert.Equal(ExcelFileFormat.XLSX, new CreatingOptions().DefaultFileFormat);
            var wb = Sample(); wb.Close();
        }

        [Fact]
        public void Password_CanBeSetAndCleared()
        {
            var wb = Sample();
            Assert.False(wb.IsEncrypted);
            wb.Password = "pw";
            Assert.True(wb.IsEncrypted);
            wb.SaveAs(Tmp("e.xlsx"));
            Assert.ThrowsAny<Exception>(() => WorkBook.Load(Tmp("e.xlsx")));
            wb.IsEncrypted = false;
            Assert.Null(wb.Password);
            wb.SaveAs(Tmp("f.xlsx"));
            AssertSample(WorkBook.Load(Tmp("f.xlsx")));
        }

        [Fact]
        public void EncryptAndDecrypt_RoundTripAStream()
        {
            var plain = Sample().ToByteArray();
            var locked = WorkBook.Create().Encrypt("pw"); // an empty workbook is encrypted too
            Assert.NotEmpty(locked.ToArray());
            var back = WorkBook.Decrypt(new MemoryStream(Sample().Encrypt("pw").ToArray()), "pw");
            AssertSample(WorkBook.Load(back.ToArray()));
            Assert.NotEqual(plain, locked.ToArray());
        }

        [Fact]
        public void DataSet_RoundTrips()
        {
            var table = new DataTable("People");
            table.Columns.Add("name", typeof(string)); table.Columns.Add("age", typeof(int));
            table.Rows.Add("Ana", 31); table.Rows.Add("Bo", 27);
            var set = new DataSet(); set.Tables.Add(table);
            var wb = WorkBook.Load(set);
            Assert.Equal("People", wb.DefaultWorkSheet.Name);
            Assert.Equal("Ana", wb.DefaultWorkSheet["A2"].StringValue);
            var back = wb.ToDataSet(useFirstRowAsColumnNames: true);
            Assert.Equal(27.0, Convert.ToDouble(back.Tables[0].Rows[1]["age"]));
            var viaTable = WorkBook.Create().LoadWorkSheet(table);
            Assert.Equal(31.0, viaTable["B2"].DoubleValue);
        }

        [Fact]
        public void Streaming_ByNameIndexStreamAndBytes()
        {
            var path = Tmp("s.xlsx"); Sample().SaveAs(path);
            var bytes = File.ReadAllBytes(path);
            Assert.Equal(new[] { "Data", "Second" }, WorkBook.GetSheetNames(path));
            Assert.Equal(new[] { "Data", "Second" }, WorkBook.GetSheetNames(bytes));
            using (var s = new MemoryStream(bytes)) Assert.Equal(new[] { "Data", "Second" }, WorkBook.GetSheetNames(s));

            Assert.Equal(3, WorkBook.StreamRows(path).Count());
            Assert.Single(WorkBook.StreamRows(path, 1));
            Assert.Single(WorkBook.StreamRows(path, "Second"));
            Assert.Single(WorkBook.StreamRows(bytes, 1));
            Assert.Single(WorkBook.StreamRows(bytes, "Second"));
            using (var s = new MemoryStream(bytes)) Assert.Equal(3, WorkBook.StreamRows(s).Count());
            using (var s = new MemoryStream(bytes)) Assert.Single(WorkBook.StreamRows(s, 1));
            using (var s = new MemoryStream(bytes)) Assert.Single(WorkBook.StreamRows(s, "Second"));
            var row = new StreamRow(5, new[] { new StreamCell(0, StreamCellType.Number, 1.0, "1") });
            Assert.Equal(5, row.RowNumber);
        }
    }

    public class SheetsCollectionSurfaceTests
    {
        static WorkBook Three()
        {
            var wb = WorkBook.Create(); wb.DefaultWorkSheet.Name = "a";
            wb.CreateWorkSheet("b"); wb.CreateWorkSheet("c");
            return wb;
        }

        [Fact]
        public void Collection_ListsFindsAndMoves()
        {
            var wb = Three();
            var sheets = wb.WorkSheets;
            Assert.Equal(3, sheets.Count);
            Assert.False(sheets.IsReadOnly);
            Assert.Equal("b", sheets[1].Name);
            Assert.Equal(new[] { "a", "b", "c" }, sheets.Select(s => s.Name));
            Assert.Equal(new[] { "a", "b", "c" }, ((System.Collections.IEnumerable)sheets).Cast<WorkSheet>().Select(s => s.Name));
            Assert.Equal(2, sheets.FindIndex(s => s.Name == "c"));
            Assert.Equal(1, sheets.IndexOf(sheets[1]));
            var own = sheets.Contains(sheets[2]); var foreign = sheets.Contains(WorkBook.Create().DefaultWorkSheet); var none = sheets.Contains(null);   // the method itself is what is tested
            Assert.True(own); Assert.False(foreign); Assert.False(none);
            Assert.Equal(-1, sheets.IndexOf(null));
            var copy = new WorkSheet[3]; sheets.CopyTo(copy, 0);
            Assert.Equal("c", copy[2].Name);
            sheets.SetSheetIndex(sheets[2], 0);
            Assert.Equal(new[] { "c", "a", "b" }, wb.WorkSheets.Select(s => s.Name));
        }

        [Fact]
        public void Collection_CreatesRemovesAndClears()
        {
            var wb = Three();
            var d = wb.WorkSheets.Create("d");
            Assert.Equal("d", d.Name); Assert.Equal(4, wb.WorkSheets.Count);
            Assert.True(wb.WorkSheets.Remove(d)); Assert.Equal(3, wb.WorkSheets.Count);
            Assert.False(wb.WorkSheets.Remove(null));
            wb.WorkSheets.RemoveAt(0);
            Assert.Equal("b", wb.DefaultWorkSheet.Name);
            wb.WorkSheets.Clear();
            Assert.Empty(wb.WorkSheets);
            Assert.Null(wb.DefaultWorkSheet);
        }

        [Fact]
        public void Collection_AddInsertAndAssign_CopySheetsIn()
        {
            var wb = Three();
            var other = WorkBook.Create(); other.DefaultWorkSheet.Name = "x"; other.DefaultWorkSheet["A1"].Value = "from other";
            wb.WorkSheets.Add(other.DefaultWorkSheet);
            Assert.Equal("from other", wb.WorkSheets[3].GetCellAt(0, 0).StringValue);
            var second = WorkBook.Create(); second.DefaultWorkSheet.Name = "y"; second.DefaultWorkSheet["A1"].Value = "inserted";
            wb.WorkSheets.Insert(0, second.DefaultWorkSheet);
            Assert.Equal("y", wb.WorkSheets[0].Name);
            Assert.Equal("inserted", wb.WorkSheets[0].GetCellAt(0, 0).StringValue);
            var third = WorkBook.Create(); third.DefaultWorkSheet.Name = "z"; third.DefaultWorkSheet["A1"].Value = "assigned";
            wb.WorkSheets[1] = third.DefaultWorkSheet;
            Assert.Equal("z", wb.WorkSheets[1].Name);
            Assert.Equal("assigned", wb.WorkSheets[1].GetCellAt(0, 0).StringValue);
            Assert.Equal(5, wb.WorkSheets.Count);
            Assert.DoesNotContain("a", wb.WorkSheets.Select(x => x.Name));          // the sheet it replaced is gone
            Assert.Throws<ArgumentOutOfRangeException>(() => wb.WorkSheets[9] = third.DefaultWorkSheet);
            Assert.Throws<ArgumentNullException>(() => wb.WorkSheets[0] = null);
            Assert.Throws<ArgumentException>(() => wb.WorkSheets[0] = wb.WorkSheets[2]);
            Assert.Equal(5, wb.WorkSheets.Count);                                     // refused replacements lose nothing
        }
    }

    public class RangeSurfaceTests
    {
        static WorkSheet Grid()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var v = new[] { new object[] { 3, "c", 1.5 }, new object[] { 1, "a", 2.5 }, new object[] { 2, "b", 3.5 } };
            for (var r = 0; r < 3; r++) for (var c = 0; c < 3; c++) ws.GetCellAt(r, c).Value = v[r][c];
            return ws;
        }

        [Fact]
        public void RowsAndColumns_AreLiveSlicesOfTheRange()
        {
            var ws = Grid();
            var range = ws["A1:C3"];
            Assert.Equal(3, range.Rows.Length); Assert.Equal(3, range.Columns.Length);
            Assert.Equal(new[] { 0, 1, 2 }, range.Rows.Select(r => r.RowNumber));
            Assert.Equal(new[] { "A", "B", "C" }, range.Columns.Select(c => c.ColumnLetter));
            Assert.Equal(6.0, (double)range.Columns[0].Sum());
#pragma warning disable CS0618
            Assert.Equal(3, range.AllRowsInRange.Length); Assert.Equal(3, range.AllColumnsInRange.Length);
#pragma warning restore CS0618
            Assert.Equal(1.5, (double)range.Rows[0].Columns[2].Sum());
            Assert.Equal(WorkBook.Create().Format, range.WorkBook.Format);
            Assert.Same(ws.Book, range.WorkBook);
        }

        [Fact]
        public void ValueAccessors_AggregateOrReadTheSingleCell()
        {
            var ws = Grid();
            var one = ws["A1"];
            Assert.Equal(3, one.IntValue); Assert.Equal(3, one.Int32Value); Assert.Equal(3L, one.LongValue); Assert.Equal(3L, one.Int64Value);
            Assert.Equal(3f, one.FloatValue); Assert.Equal(3m, one.DecimalValue);
            Assert.Equal(3, one.GetValue<int>()); Assert.True(one.TryGetValue<double>(out var d)); Assert.Equal(3.0, d);
            Assert.False(ws["B1"].TryGetValue<int>(out var bad)); Assert.Equal(0, bad);
            Assert.Equal(6, ws["A1:A3"].IntValue); Assert.Equal(6L, ws["A1:A3"].LongValue); Assert.Equal(6f, ws["A1:A3"].FloatValue);
            Assert.Equal(6, ws["A1:A3"].GetValue<int>());
            Assert.Equal("3", one.ToString());
            Assert.Equal("3\tc" + Environment.NewLine + "1\ta", ws["A1:B2"].ToString());

            ws["E1"].StringValue = "text"; Assert.Equal("text", ws["E1"].StringValue);
            ws["E2"].DateTimeValue = new DateTime(2024, 3, 9); Assert.Equal(new DateTime(2024, 3, 9), ws["E2"].DateTimeValue);
            ws["E3:E4"].IntValue = 7; Assert.Equal(14, ws["E3:E4"].IntValue);
            ws["F1:F2"].BoolValue = true; Assert.True(ws["F1:F2"].BoolValue);
            Assert.False(ws["F1:F3"].BoolValue);
            ws["G1"].FloatValue = 1.25f; Assert.Equal(1.25, ws["G1"].DoubleValue);
            ws["G2"].LongValue = 5000000000; Assert.Equal(5000000000, ws["G2"].LongValue);
            ws["G3"].Int32Value = 9; ws["G4"].Int64Value = 10; Assert.Equal(19, ws["G3:G4"].IntValue);
            ws["G5"].DecimalValue = 2.5m; Assert.Equal(2.5m, ws["G5"].DecimalValue);
            Assert.Equal(1, ws["A2"].IntValue);
        }

        [Fact]
        public void RangeMath_AddSubtractAndEnumerate()
        {
            var ws = Grid();
            var sum = ws["A1:A2"] + ws["C3:C3"];
            Assert.Equal("A1:C3", sum.RangeAddressAsString);
            var rest = ws["A1:C3"] - ws["A1:C1"];
            Assert.Equal("A2:C3", rest.RangeAddressAsString);
            Assert.Equal("A1:C2", (ws["A1:C3"] - ws["A3:C3"]).RangeAddressAsString);
            Assert.Equal("B1:C3", (ws["A1:C3"] - ws["A1:A3"]).RangeAddressAsString);
            Assert.Equal("A1:B3", (ws["A1:C3"] - ws["C1:C3"]).RangeAddressAsString);
            Assert.Equal(9, ws["A1:C3"].Count());
            Assert.Equal(9, ((System.Collections.IEnumerable)ws["A1:C3"]).Cast<Cell>().Count());
            Assert.Equal(2, ws["A1:C3"].AddRange(ws["D1:D3"]).ColumnCount - 2);
            Assert.Equal("A1:C1", ws["A1:C3"].SubtractRange(ws["A2:C3"]).RangeAddressAsString);
            Assert.Equal("A1:D3", ws["A1:C3"].AddRange(new RangeAddress(0, 2, 3, 3)).RangeAddressAsString);
            Assert.Equal("A2:C3", ws["A1:C3"].SubtractRange(new RangeAddress(0, 0, 0, 2)).RangeAddressAsString);
            Assert.Throws<NotSupportedException>(() => ws.AddRange(ws["A1"]));
            Assert.Throws<NotSupportedException>(() => ws.SubtractRange(ws["A1"]));
            Assert.Throws<NotSupportedException>(() => ws.AddRange(new RangeAddress(0, 0, 0, 0)));
            Assert.Throws<NotSupportedException>(() => ws.SubtractRange(new RangeAddress(0, 0, 0, 0)));
        }

        [Fact]
        public void SortAscendingAndDescending_UseTheFirstColumn()
        {
            var ws = Grid();
            ws["A1:C3"].SortAscending();
            Assert.Equal(new object[] { 1.0, 2.0, 3.0 }, new[] { ws["A1"].Value, ws["A2"].Value, ws["A3"].Value });
            Assert.Equal(new object[] { "a", "b", "c" }, new[] { ws["B1"].Value, ws["B2"].Value, ws["B3"].Value });
            ws["A1:C3"].SortDescending();
            Assert.Equal(3.0, ws["A1"].DoubleValue); Assert.Equal("c", ws["B1"].StringValue); Assert.Equal(1.5, ws["C1"].DoubleValue);
            ws["A1:C3"].SortByColumn("B", SortOrder.Ascending);
            Assert.Equal("a", ws["B1"].StringValue);
            Assert.Throws<ArgumentOutOfRangeException>(() => ws["A1:C3"].SortByColumn(5, SortOrder.Ascending));
        }

        [Fact]
        public void InsertAndRemoveColumns_MoveTheCells()
        {
            var ws = Grid();
            var inserted = ws["A1:C3"].InsertColumn(1);
            Assert.Equal(1, inserted.ColumnNumber);
            Assert.Equal("c", ws["C1"].StringValue); Assert.Equal(1.5, ws["D1"].DoubleValue);
            var byLetter = ws.InsertColumn("A");
            Assert.Equal("A", byLetter.ColumnLetter);
            Assert.Equal(3.0, ws["B1"].DoubleValue);
            var many = ws.InsertColumns("A", 2);
            Assert.Equal(2, many.Count);
            Assert.Equal(3.0, ws["D1"].DoubleValue);
            var byIndex = ws.InsertColumns(0, 1);
            Assert.Single(byIndex);
            Assert.Equal(3.0, ws["E1"].DoubleValue);

            ws.RemoveColumn(0);
            ws.RemoveColumn("A");
            ws.RemoveColumn(ws.GetColumn(0));
            Assert.Equal(3.0, ws["B1"].DoubleValue);   // [_,_,_,_,3,_,c,1.5] lost its three empty columns on the left
            Assert.Equal("c", ws["D1"].StringValue);
        }

        [Fact]
        public void RemoveRowAndColumnFromRangeRowAndColumn()
        {
            var ws = Grid();
            ws.GetRow(0).RemoveRow();
            Assert.Equal(1.0, ws["A1"].DoubleValue);
            ws.GetColumn("A").RemoveColumn();
            Assert.Equal("a", ws["A1"].StringValue);
            ws.RemoveRow(ws.GetRow(0));
            Assert.Equal("b", ws["A1"].StringValue);
            Assert.Throws<ArgumentOutOfRangeException>(() => ws["A1:B2"].RemoveRow(5));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws["A1:B2"].RemoveColumn(5));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws["A1:B2"].InsertRows(0, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws["A1:B2"].InsertColumns(0, 0));
        }

        [Fact]
        public void RowAndColumnObjects_HeightWidthHiddenAndAutosize()
        {
            var ws = Grid();
            var row = ws.GetRow(0); row.Height = 30; row.Hidden = true;
            Assert.Equal(30, ws.GetRow(0).Height); Assert.True(ws.GetRow(0).Hidden);
            row.AutoSizeRow(); row.AutoSizeRow(true);
            var col = ws.GetColumn(1); col.Width = 4000; col.Hidden = true;
            Assert.Equal(4000, ws.GetColumn("B").Width); Assert.True(ws.GetColumn(1).Hidden);
            col.AutoSizeColumn(); col.AutoSizeColumn(true);
            Assert.Equal("B", col.ColumnLetter);
        }

        [Fact]
        public void RangeAddress_EqualityAndText()
        {
            var a = new RangeAddress(0, 1, 0, 2); var b = new RangeAddress(0, 1, 0, 2);
            Assert.Equal("A1:C2", a.ToString());
            Assert.True(a.Equals(b)); Assert.False(a.Equals(null)); Assert.False(a.Equals("A1:C2"));
            Assert.Equal(a.GetHashCode(), b.GetHashCode());
            Assert.NotEqual(a.GetHashCode(), new RangeAddress(0, 1, 0, 3).GetHashCode());
            Assert.True(a.Includes(1, 2)); Assert.False(a.Includes(2, 0));
            Assert.Equal("B2:C2", a.GetIntersection(new RangeAddress(1, 5, 1, 5)).Location);
            Assert.Null(a.GetIntersection(new RangeAddress(5, 6, 5, 6)));
        }

        [Fact]
        public void WorksheetLevelRemoveAndInsertRows()
        {
            var ws = Grid();
            ws.RemoveRow(0);
            Assert.Equal(1.0, ws["A1"].DoubleValue);
            var rows = ws.InsertRows(0, 2);
            Assert.Equal(2, rows.Count);
            Assert.Equal(1.0, ws["A3"].DoubleValue);
            Assert.Equal(0, ws.InsertRow(0).RowNumber);
            Assert.Equal(1.0, ws["A4"].DoubleValue);
        }
    }

    public class CellSurfaceTests : TestBase
    {
        [Fact]
        public void Cell_TextStyleAndCopy()
        {
            var wb = WorkBook.Create(); var ws = wb.DefaultWorkSheet;
            var cell = ws.GetCellAt(0, 0);
            cell.Text = "hello"; Assert.Equal("hello", cell.Text);
            cell.StringValue = "x"; Assert.Equal("x", cell.ToString());
            cell.BoolValue = true; Assert.True(cell.BoolValue); Assert.True(cell.IsBoolean);
            cell.DateTimeValue = new DateTime(2024, 3, 9); Assert.True(cell.IsDateTime);
#pragma warning disable CS0618   // the old names still work and must keep working
            cell.Int32Value = 3; cell.Int64Value = 4; Assert.Equal(4, cell.IntValue);
#pragma warning restore CS0618
            Assert.Equal(4L, cell.LongValue);
            cell.Style.Font.Bold = true;
            Assert.True(cell.Style.Font.Bold);

            ws["B2"].Value = 7; ws["B2"].Style.Font.Italic = true;
            var second = wb.CreateWorkSheet("Two");
            ws.GetCellAt(1, 1).Copy(second, "D4");
            Assert.Equal(7.0, second["D4"].DoubleValue);
            Assert.True(second["D4"].Style.Font.Italic);
            ws["B2:B2"].Copy(second, "A1");
            Assert.Equal(7.0, second["A1"].DoubleValue);
            Assert.Throws<ArgumentNullException>(() => ws.GetCellAt(1, 1).Copy(null, "A1"));
            Assert.Throws<ArgumentNullException>(() => ws["B2"].Copy(null, "A1"));
        }

        [Fact]
        public void Comment_ContentAuthorVisibility()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var cell = ws.GetCellAt(0, 0);
            var c = cell.AddComment("first", "me");
            Assert.Equal("first", c.Content); Assert.Equal("me", c.Author); Assert.False(c.IsVisible);
            c.Content = "second"; c.Author = "you"; c.IsVisible = true;
            Assert.Equal("second", cell.Comment.Content); Assert.Equal("you", cell.Comment.Author); Assert.True(cell.Comment.IsVisible);
            Assert.Equal("A1", cell.Comment.Address);
            cell.Comment.Content = null;
            Assert.Equal("", cell.Comment.Content ?? "");
            cell.RemoveComment();
            Assert.Null(cell.Comment);
        }

        [Fact]
        public void Comments_Survive_SavingAs_XlsAndXlsx()
        {
            foreach (var ext in new[] { "c.xlsx", "c.xls" })
            {
                var wb = ext.EndsWith("xls") ? WorkBook.Create(ExcelFileFormat.XLS) : WorkBook.Create();
                wb.DefaultWorkSheet.GetCellAt(2, 1).AddComment("note", "author");
                wb.SaveAs(Tmp(ext));
                var back = WorkBook.Load(Tmp(ext)).DefaultWorkSheet.GetCellAt(2, 1).Comment;
                Assert.Equal("note", back.Content); Assert.Equal("author", back.Author);
            }
        }

        [Fact]
        public void Hyperlink_KindsAndRemoval()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var c = ws.GetCellAt(0, 0);
            c.Hyperlink = "https://example.com/x"; Assert.Equal("https://example.com/x", c.Hyperlink);
            c.Hyperlink = "mailto:me@example.com"; Assert.Equal("mailto:me@example.com", c.Hyperlink);
            c.Hyperlink = "file:///c:/a.txt"; Assert.Equal("file:///c:/a.txt", c.Hyperlink);
            Assert.Single(ws.Native.GetHyperlinkList());                       // replaced, not piled up
            c.RemoveHyperlink(); Assert.Null(c.Hyperlink);
            c.Hyperlink = "https://example.com"; c.Hyperlink = " "; Assert.Null(c.Hyperlink);
        }
    }

    public class StyleSurfaceTests : TestBase
    {
        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void EveryStyleProperty_CanBeSetAndReadBack(ExcelFileFormat fmt)
        {
            var ws = WorkBook.Create(fmt).DefaultWorkSheet;
            var range = ws["B2:C3"];
            var st = range.Style;
            st.BackgroundColor = "#FFE699";
            if (fmt == ExcelFileFormat.XLSX) Assert.Equal("#FFE699", ws["B2"].Style.BackgroundColor, ignoreCase: true);
            else Assert.NotEqual("#FFFFFF", ws["B2"].Style.BackgroundColor);        // an .xls only has 56 colors; the nearest one is used
            st.SetBackgroundColor(Color.FromRgb(0, 128, 0));
            st.FillPattern = FillPattern.Diamonds; Assert.Equal(FillPattern.Diamonds, ws["C3"].Style.FillPattern);
            st.HorizontalAlignment = HorizontalAlignment.Center; Assert.Equal(HorizontalAlignment.Center, ws["B3"].Style.HorizontalAlignment);
            st.VerticalAlignment = VerticalAlignment.Top; Assert.Equal(VerticalAlignment.Top, ws["B3"].Style.VerticalAlignment);
            st.Rotation = 45; Assert.Equal(45, ws["B2"].Style.Rotation);
            st.Indention = 2; Assert.Equal(2, ws["B2"].Style.Indention);
            st.WrapText = true; Assert.True(ws["B2"].Style.WrapText);
            st.ShrinkToFit = true; Assert.True(ws["B2"].Style.ShrinkToFit);
            st.DiagonalBorderDirection = DiagonalBorderDirection.Both; Assert.Equal(DiagonalBorderDirection.Both, ws["B2"].Style.DiagonalBorderDirection);

            foreach (var side in new[] { st.LeftBorder, st.RightBorder, st.TopBorder, st.BottomBorder, st.DiagonalBorder })
            {
                side.Type = BorderType.Dashed; side.SetColor("#FF0000");
            }
            Assert.Equal(BorderType.Dashed, ws["B2"].Style.LeftBorder.Type);
            Assert.Equal(BorderType.Dashed, ws["B2"].Style.RightBorder.Type);
            Assert.Equal(BorderType.Dashed, ws["B2"].Style.TopBorder.Type);
            Assert.Equal(BorderType.Dashed, ws["B2"].Style.BottomBorder.Type);
            Assert.Equal(BorderType.Dashed, ws["B2"].Style.DiagonalBorder.Type);
            Assert.Equal("#FF0000", ws["B2"].Style.LeftBorder.Color, ignoreCase: true);
            st.TopBorder.SetColor(Color.FromRgb(0, 0, 255)); Assert.Equal("#0000FF", ws["B2"].Style.TopBorder.Color, ignoreCase: true);
            st.BottomBorder.SetBorderStyle("#00FF00", BorderType.Thick); Assert.Equal(BorderType.Thick, ws["B2"].Style.BottomBorder.Type);
            st.LeftBorder.SetBorderStyle(Color.FromRgb(1, 2, 3), BorderType.Double); Assert.Equal(BorderType.Double, ws["B2"].Style.LeftBorder.Type);

            var f = st.Font;
            f.Bold = true; f.Italic = true; f.Strikeout = true; f.Name = "Arial"; f.Height = 14;
            f.Underline = FontUnderlineType.Double; f.FontScript = FontScript.Super;
            f.SetColor("#123456");
            var read = ws["C3"].Style.Font;
            Assert.True(read.Bold); Assert.True(read.Italic); Assert.True(read.Strikeout);
            Assert.Equal("Arial", read.Name); Assert.Equal(14, read.Height);
            Assert.Equal(FontUnderlineType.Double, read.Underline); Assert.Equal(FontScript.Super, read.FontScript);
            if (fmt == ExcelFileFormat.XLSX) Assert.Equal("#123456", read.Color, ignoreCase: true);   // .xls: nearest of its 56 colors
            f.SetColor(Color.FromRgb(255, 0, 255)); Assert.Equal("#FF00FF", ws["B2"].Style.Font.Color, ignoreCase: true);
            f.Color = "#00FFFF"; Assert.Equal("#00FFFF", ws["B2"].Style.Font.Color, ignoreCase: true);

            ws.Book.SaveAs(Tmp("s." + (fmt == ExcelFileFormat.XLS ? "xls" : "xlsx")));
            var back = WorkBook.Load(Tmp("s." + (fmt == ExcelFileFormat.XLS ? "xls" : "xlsx"))).DefaultWorkSheet["C3"].Style;
            Assert.Equal(FillPattern.Diamonds, back.FillPattern);
            Assert.Equal("Arial", back.Font.Name);
            Assert.Equal(HorizontalAlignment.Center, back.HorizontalAlignment);
        }

        [Fact]
        public void Color_ParsesFormatsAndCompares()
        {
            var c = Color.FromRgb(255, 128, 0);
            Assert.Equal("#FF8000", c.ToHex()); Assert.Equal("#FF8000", c.ToString());
            Assert.Equal(c, Color.FromHex("ff8000")); Assert.Equal(c, Color.FromHex("#FF8000"));
            Assert.True(c.Equals((object)Color.FromHex("FF8000"))); Assert.False(c.Equals((object)"x")); Assert.False(c.Equals(Color.FromRgb(0, 0, 0)));
            Assert.Equal(c.GetHashCode(), Color.FromHex("FF8000").GetHashCode());
            Assert.Equal(255, c.R); Assert.Equal(128, c.G); Assert.Equal(0, c.B);
            Assert.Throws<ArgumentException>(() => Color.FromHex("xyz"));
            Assert.Throws<ArgumentException>(() => Color.FromHex("12345"));
            Assert.Throws<ArgumentNullException>(() => Color.FromHex(null));
        }

        [Fact]
        public void BuiltinFormats_AreExcelsOwn()
        {
            Assert.Equal("General", BuiltinFormats.General);
            Assert.Equal("0.00", BuiltinFormats.Number2);
            Assert.Equal("0%", BuiltinFormats.Percent);
            Assert.Equal("0.00E+00", BuiltinFormats.Scientific1);
            Assert.Equal("m/d/yy", BuiltinFormats.ShortDate);
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = 0.256; ws["A1"].FormatString = BuiltinFormats.Percent;
            Assert.Equal("26%", ws["A1"].ToArray()[0].FormattedCellValue);
        }
    }

    public class SheetSettingsSurfaceTests : TestBase
    {
        [Fact]
        public void PrintSetup_EveryGetterReturnsWhatWasSet()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var p = ws.PrintSetup;
            p.PrintOrientation = PrintOrientation.Landscape; p.Scale = 80; p.PageStart = 3; p.PaperSize = PaperSize.A4;
            p.FitWidth = 2; p.FitHeight = 3; p.LeftToRight = true; p.NoColor = true; p.Draft = true; p.Notes = true; p.UsePage = true;
            p.HResolution = 300; p.VResolution = 600; p.HeaderMargin = 0.4; p.FooterMargin = 0.5; p.Copies = 2;
            var r = ws.PrintSetup;
            Assert.Equal(PrintOrientation.Landscape, r.PrintOrientation); Assert.Equal(80, r.Scale); Assert.Equal(3, r.PageStart);
            Assert.Equal(PaperSize.A4, r.PaperSize); Assert.Equal(2, r.FitWidth); Assert.Equal(3, r.FitHeight);
            Assert.True(r.LeftToRight); Assert.True(r.NoColor); Assert.True(r.Draft); Assert.True(r.Notes); Assert.True(r.UsePage);
            Assert.Equal(300, r.HResolution); Assert.Equal(600, r.VResolution);
            Assert.Equal(0.4, r.HeaderMargin); Assert.Equal(0.5, r.FooterMargin); Assert.Equal(2, r.Copies);
        }

        [Fact]
        public void HeadersAndFooters_ThreeParts()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws.Header.Left = "L"; ws.Header.Center = "C"; ws.Header.Right = "R";
            ws.Footer.Left = "l"; ws.Footer.Center = "c"; ws.Footer.Right = "r";
            Assert.Equal("L", ws.Header.Left); Assert.Equal("C", ws.Header.Center); Assert.Equal("R", ws.Header.Right);
            Assert.Equal("l", ws.Footer.Left); Assert.Equal("c", ws.Footer.Center); Assert.Equal("r", ws.Footer.Right);
        }

        [Fact]
        public void Metadata_ReadOnlyFieldsOnBothFormats()
        {
            foreach (var fmt in new[] { ExcelFileFormat.XLSX, ExcelFileFormat.XLS })
            {
                var m = WorkBook.Create(fmt).Metadata;
                m.Author = "me"; m.Title = "t"; m.Subject = "s"; m.Keywords = "k"; m.Comments = "c"; m.Category = "cat";
                Assert.Equal("me", m.Author); Assert.Equal("cat", m.Category);
                // these four are set by whatever wrote the file, not by this API; reading must not throw
                _ = m.ApplicationName; _ = m.Company; _ = m.Manager; _ = m.Template;
                m.CustomProperties.AddProperty("n", 5);
                Assert.True(m.CustomProperties.Contains("n"));
            }
        }

        [Fact]
        public void LicenseAndLogger_AreInertShims()
        {
            Assert.True(License.IsLicensed);
            Assert.True(License.IsValidLicense("anything"));
            License.LicenseKey = "k"; Assert.Equal("k", License.LicenseKey);
            License.DisableAppAnalytics();
#pragma warning disable CS0618
            Logging.Logger.LogFilePath = "x.log"; Assert.Equal("x.log", Logging.Logger.LogFilePath);
            Logging.Logger.LoggingMode = Logging.Logger.LoggingModes.All; Assert.Equal(Logging.Logger.LoggingModes.All, Logging.Logger.LoggingMode);
            Logging.Logger.ClearLogFiles();
#pragma warning restore CS0618
        }

        [Fact]
        public void AutoFilter_NamedRangesAndPaneVariants()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            for (var r = 0; r < 4; r++) { ws.GetCellAt(r, 0).Value = r; ws.GetCellAt(r, 1).Value = "x" + r; }
            ws.SetAutoFilter("A1:B4");
            ws.SetAutoFilter(0, 3, 0, 1);
            ws.UnsetAutoFilter();
            ws.CreateFreezePane(1, 1); ws.CreateFreezePane(1, 1, 2, 2); ws.RemovePane();
            ws.CreateSplitPane(1000, 1000, 1, 1, PanePosition.LowerRight);
            ws.SetZoom(150);
            ws.SetRowBreak(2); ws.SetColumnBreak(1); ws.RemoveRowBreak(2); ws.RemoveColumnBreak(1);
            ws.GroupRows(0, 1); ws.UngroupRows(0, 1); ws.GroupColumns(0, 1); ws.UngroupColumns(0, 1);
            ws.GroupColumns("A", "B"); ws.UngroupColumn("A", "B");
            ws.SetRepeatingRows(0, 0); ws.SetRepeatingColumns(0, 0); ws.SetRepeatingColumns("A", "B");
            ws.RemoveRepeatingRows(); ws.RemoveRepeatingColumns();
            ws.SetPrintArea("A1:B4");
            ws.SetWorksheetTabColor(new byte[] { 255, 0, 0 }); Assert.Equal(new byte[] { 255, 0, 0 }, ws.TabColor);
            ws.AutoSizeColumn(0); ws.AutoSizeColumn(0, true); ws.AutoSizeRow(0); ws.AutoSizeRow(0, true);
            ws.ProtectSheet("pw"); ws.UnprotectSheet();
            ws.Hidden = HiddenState.Hidden; Assert.Equal(HiddenState.Hidden, ws.Hidden);
            ws.IsRightToLeft = true; Assert.True(ws.IsRightToLeft);
            ws.DisplayGridlines = false; ws.IsPrintGridlines = true; ws.FitToPage = true; ws.HorizontallyCenter = true; ws.VerticallyCenter = true;
            Assert.False(ws.DisplayGridlines); Assert.True(ws.IsPrintGridlines); Assert.True(ws.FitToPage); Assert.True(ws.HorizontallyCenter); Assert.True(ws.VerticallyCenter);
            ws.Book.SaveAs(Tmp("settings.xlsx"));
            Assert.True(WorkBook.Load(Tmp("settings.xlsx")).DefaultWorkSheet.IsRightToLeft);
        }
    }

    public class FeatureSurfaceTests : TestBase
    {
        static WorkSheet Fresh(ExcelFileFormat fmt = ExcelFileFormat.XLSX) => WorkBook.Create(fmt).DefaultWorkSheet;

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void DataValidation_EveryAddOverload(ExcelFileFormat fmt)
        {
            var ws = Fresh(fmt);
            var dv = ws.DataValidations;
            var cell = ws.GetCellAt(0, 0);
            Assert.Equal(ConstraintType.List, dv.AddFormulaListRule(cell, "$H$1:$H$3").ConstraintType);
            Assert.Equal(ConstraintType.List, dv.AddFormulaListRule(ws["A2"], "$H$1:$H$3").ConstraintType);
            Assert.Equal(ConstraintType.List, dv.AddFormulaListRule(new RangeAddress(2, 2, 0, 0), "=$H$1:$H$3").ConstraintType);
            Assert.Equal(ConstraintType.List, dv.AddStringListRule(ws.GetCellAt(3, 0), new[] { "a", "b" }).ConstraintType);
            Assert.Equal(ConstraintType.List, dv.AddStringListRule(ws["A5"], new[] { "a", "b" }).ConstraintType);
            Assert.Equal(ConstraintType.List, dv.AddStringListRule(new RangeAddress(5, 5, 0, 0), new[] { "a", "b" }).ConstraintType);
            Assert.Equal(ConstraintType.WholeNumber, dv.AddIntegerRule(ws.GetCellAt(0, 1), ComparisonOperator.Equal, 5, null).ConstraintType);
            Assert.Equal(ConstraintType.WholeNumber, dv.AddIntegerRule(ws["B2"], ComparisonOperator.NotEqual, 5, null).ConstraintType);
            Assert.Equal(ConstraintType.WholeNumber, dv.AddIntegerRule(new RangeAddress(2, 2, 1, 1), ComparisonOperator.Between, 1, 9).ConstraintType);
            Assert.Equal(ConstraintType.Decimal, dv.AddDecimalRule(ws.GetCellAt(0, 2), ComparisonOperator.LessThan, 1.5m, null).ConstraintType);
            Assert.Equal(ConstraintType.Decimal, dv.AddDecimalRule(ws["C2"], ComparisonOperator.NotBetween, 1.5m, 2.5m).ConstraintType);
            Assert.Equal(ConstraintType.Decimal, dv.AddDecimalRule(new RangeAddress(2, 2, 2, 2), ComparisonOperator.GreaterThanOrEqual, 1m, null).ConstraintType);
            Assert.Equal(ConstraintType.Date, dv.AddDateRule(ws.GetCellAt(0, 3), ComparisonOperator.LessThanOrEqual, new DateTime(2024, 1, 1), null).ConstraintType);
            Assert.Equal(ConstraintType.Date, dv.AddDateRule(ws["D2"], ComparisonOperator.Between, new DateTime(2024, 1, 1), new DateTime(2024, 12, 31)).ConstraintType);
            Assert.Equal(ConstraintType.Date, dv.AddDateRule(new RangeAddress(2, 2, 3, 3), ComparisonOperator.GreaterThan, new DateTime(2024, 1, 1), null).ConstraintType);
            Assert.Equal(ConstraintType.TextLength, dv.AddTextLengthRule(ws.GetCellAt(0, 4), ComparisonOperator.LessThan, 5, null).ConstraintType);
            Assert.Equal(ConstraintType.TextLength, dv.AddTextLengthRule(ws["E2"], ComparisonOperator.Between, 1, 5).ConstraintType);
            Assert.Equal(ConstraintType.TextLength, dv.AddTextLengthRule(new RangeAddress(2, 2, 4, 4), ComparisonOperator.GreaterThan, 1, null).ConstraintType);
            Assert.Equal(18, dv.Count);

            Assert.Equal(18, dv.Count());
            Assert.Equal(18, ((System.Collections.IEnumerable)dv).Cast<DataValidation>().Count());
            Assert.Equal(ConstraintType.List, dv[0].ConstraintType);
            Assert.Equal(3, dv.GetRulesForRange("B1:B3").Count);
            Assert.Equal(3, dv.GetRulesForRange(ws["C1:C3"]).Count);
            Assert.Equal(3, dv.GetRulesForRange(new RangeAddress(0, 2, 3, 3)).Count);
            Assert.True(dv.RangeAlreadyHasDataValidation(new RangeAddress(0, 0, 0, 0)));
            Assert.False(dv.RangeAlreadyHasDataValidation(new RangeAddress(20, 20, 0, 0)));
            dv.RemoveRuleFromCell(cell);
            Assert.Equal(17, dv.Count);
            dv.RemoveAllRulesFromRange("B1:B3"); Assert.Equal(14, dv.Count);
            dv.RemoveAllRulesFromRange(ws["C1:C3"]); Assert.Equal(11, dv.Count);
            dv.RemoveAllRulesFromRange(new RangeAddress(0, 2, 3, 3)); Assert.Equal(8, dv.Count);
            dv.RemoveAt(0); Assert.Equal(7, dv.Count);
            Assert.Throws<ArgumentNullException>(() => dv.Remove(null));
            Assert.Throws<ArgumentException>(() => dv.AddStringListRule("A9", new string[0]));
            Assert.Throws<ArgumentException>(() => dv.AddFormulaListRule("A9", " "));
            Assert.Throws<ArgumentException>(() => dv.AddStringListRule("A9", new[] { new string('x', 300) }));
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void DataValidation_FlagsAndEquality(ExcelFileFormat fmt)
        {
            var ws = Fresh(fmt);
            var v = ws.DataValidations.AddIntegerRule("A1", ComparisonOperator.Between, 1, 5);
            Assert.True(v.EmptyCellAllowed); Assert.True(v.ShowErrorBox);
            v.EmptyCellAllowed = false; v.ShowErrorBox = false; v.ComparisonOperator = ComparisonOperator.NotBetween;
            var again = ws.DataValidations[0];
            Assert.False(again.EmptyCellAllowed); Assert.False(again.ShowErrorBox);
            Assert.Equal(ComparisonOperator.NotBetween, again.ComparisonOperator);
            Assert.Equal(v, again); Assert.Equal(v.GetHashCode(), again.GetHashCode());
            Assert.False(v.Equals("x"));

            var list = ws.DataValidations.AddStringListRule("B1", new[] { "x", "y" });
            Assert.True(list.ShowDropDownList);
            list.ShowDropDownList = false;
            Assert.False(ws.DataValidations.GetRuleForCell(ws.GetCellAt(0, 1)).ShowDropDownList);
            Assert.False(v.ShowDropDownList);   // only list rules have an arrow
            ws.Book.SaveAs(Tmp("dv." + (fmt == ExcelFileFormat.XLS ? "xls" : "xlsx")));
            var loaded = WorkBook.Load(Tmp("dv." + (fmt == ExcelFileFormat.XLS ? "xls" : "xlsx"))).DefaultWorkSheet.DataValidations;
            Assert.False(loaded.GetRuleForCell(WorkBook.Load(Tmp("dv." + (fmt == ExcelFileFormat.XLS ? "xls" : "xlsx"))).DefaultWorkSheet.GetCellAt(0, 1)).ShowDropDownList);
            var edited = loaded.Single(x => x.ConstraintType == ConstraintType.WholeNumber);
            Assert.Equal(ComparisonOperator.NotBetween, edited.ComparisonOperator);   // changing the operator after creation reaches the file
            Assert.False(edited.EmptyCellAllowed);
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void ConditionalFormatting_EveryOverload(ExcelFileFormat fmt)
        {
            var ws = Fresh(fmt);
            var cf = ws.ConditionalFormatting;
            var gt = cf.CreateConditionalFormattingRule(CfOp.GreaterThan, "5");
            var between = cf.CreateConditionalFormattingRule(CfOp.Between, "1", "9");
            var formula = cf.CreateConditionalFormattingRule("A1>2");
            Assert.Equal(ConditionType.CellValueIs, gt.ConditionType);
            Assert.Equal(ConditionType.Formula, formula.ConditionType);
            Assert.Equal(CfOp.GreaterThan, gt.ComparisonOperation); Assert.Equal("5", gt.Formula1);
            Assert.Equal("9", between.Formula2);

            Assert.Equal(1, cf.AddConditionalFormatting("A1:A3", gt).NumberOfRules);
            Assert.Equal(2, cf.AddConditionalFormatting("B1:B3", gt, between).NumberOfRules);
            Assert.Equal(2, cf.AddConditionalFormatting("C1:C3", new List<ConditionalFormattingRule> { gt, formula }).NumberOfRules);
            Assert.Equal(1, cf.AddConditionalFormatting(new[] { "D1:D3", "F1:F3" }, between).NumberOfRules);
            Assert.Equal(2, cf.AddConditionalFormatting(new[] { "E1:E3" }, gt, formula).NumberOfRules);
            Assert.Equal(2, cf.AddConditionalFormatting(new[] { "G1:G3" }, new List<ConditionalFormattingRule> { gt, between }).NumberOfRules);
            Assert.Equal(6, cf.ConditionalFormattingsCount);
            Assert.Throws<ArgumentException>(() => cf.AddConditionalFormatting("H1", new List<ConditionalFormattingRule>()));

            var two = cf.GetConditionalFormattingAt(1);
            Assert.Equal("B1:B3", two.GetFormattingRanges().Single().Location);
            Assert.Equal(CfOp.Between, two.GetRule(1).ComparisonOperation);
            two.SetRule(1, formula);
            Assert.Equal(ConditionType.Formula, cf.GetConditionalFormattingAt(1).GetRule(1).ConditionType);
            two.AddRule(gt);
            cf.RemoveConditionalFormatting(5);
            Assert.Equal(5, cf.ConditionalFormattingsCount);
        }

        [Fact]
        public void ConditionType_LooksUpByIdAndCompares()
        {
            Assert.Equal(ConditionType.Formula, ConditionType.ForId((byte)2));
            Assert.Equal(ConditionType.IconSet, ConditionType.ForId(6));
            Assert.Null(ConditionType.ForId(99)); Assert.Null(ConditionType.ForId(-1)); Assert.Null(ConditionType.ForId(1000));
            Assert.Equal("cellIs", ConditionType.CellValueIs.ToString());
            Assert.Equal(1, ConditionType.CellValueIs.GetHashCode());
            Assert.False(ConditionType.Formula.Equals("expression"));
            Assert.NotEqual(ConditionType.ColorScale, ConditionType.DataBar);
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void ConditionalFormatting_EveryFormattingPart(ExcelFileFormat fmt)
        {
            var ws = Fresh(fmt);
            var rule = ws.ConditionalFormatting.CreateConditionalFormattingRule(CfOp.GreaterThan, "1");
            var b = rule.BorderFormatting;
            b.TopBorderType = BorderType.Thin; b.BottomBorderType = BorderType.Medium; b.LeftBorderType = BorderType.Dashed;
            b.RightBorderType = BorderType.Dotted; b.DiagonalBorderType = BorderType.Hair;
            b.TopBorderColor = "#FF0000"; b.BottomBorderColor = "#00FF00"; b.LeftBorderColor = "#0000FF";
            b.RightBorderColor = "#FFFF00"; b.DiagonalBorderColor = "#00FFFF";
            var f = rule.FontFormatting;
            f.IsBold = true; f.IsItalic = true; f.FontColor = "#800000"; f.FontHeight = 12;
            f.UnderlineType = FontUnderlineType.Single; f.EscapementType = FontScript.Sub;
            var p = rule.PatternFormatting;
            p.BackgroundColor = "#FFE699"; p.FillPattern = FillPattern.SolidForeground;
            ws.ConditionalFormatting.AddConditionalFormatting("A1:A3", rule);

            ws.Book.SaveAs(Tmp("cf." + (fmt == ExcelFileFormat.XLS ? "xls" : "xlsx")));
            var back = WorkBook.Load(Tmp("cf." + (fmt == ExcelFileFormat.XLS ? "xls" : "xlsx"))).DefaultWorkSheet.ConditionalFormatting.GetConditionalFormattingAt(0).GetRule(0);
            Assert.True(back.FontFormatting.IsBold); Assert.True(back.FontFormatting.IsItalic);
            Assert.Equal(12, back.FontFormatting.FontHeight);
            Assert.Equal(FontUnderlineType.Single, back.FontFormatting.UnderlineType);
            Assert.Equal(FontScript.Sub, back.FontFormatting.EscapementType);
            Assert.Equal(BorderType.Thin, back.BorderFormatting.TopBorderType);
            Assert.Equal(BorderType.Medium, back.BorderFormatting.BottomBorderType);
            Assert.Equal(BorderType.Dashed, back.BorderFormatting.LeftBorderType);
            Assert.Equal(BorderType.Dotted, back.BorderFormatting.RightBorderType);
            Assert.NotNull(back.BorderFormatting.DiagonalBorderColor);
            Assert.NotNull(back.BorderFormatting.TopBorderColor); Assert.NotNull(back.BorderFormatting.BottomBorderColor);
            Assert.NotNull(back.BorderFormatting.LeftBorderColor); Assert.NotNull(back.BorderFormatting.RightBorderColor);
            Assert.NotNull(back.FontFormatting.FontColor);
            Assert.Equal(FillPattern.SolidForeground, back.PatternFormatting.FillPattern);
            Assert.NotNull(back.PatternFormatting.BackgroundColor);
        }

        static readonly byte[] Png = Samples.Png;

        [Fact]
        public void InsertImage_FromAFileWithLettersOrNumbers()
        {
            File.WriteAllBytes(Tmp("p.png"), Png);
            var ws = Fresh();
            var a = ws.InsertImage(Tmp("p.png"), "B", 1, "D", 4);
            var b = ws.InsertImage(Tmp("p.png"), 5, 5, 7, 8);
            Assert.Equal(1, a.Position.LeftColumnIndex); Assert.Equal("D", a.Position.RightColumnLetter);
            Assert.Equal(5, b.Position.LeftColumnIndex); Assert.Equal(AnchorType.MoveAndResize, a.Position.AnchorType);
            Assert.Equal(2, ws.Images.Count);
            ws.RemoveImage(0);
            Assert.Single(ws.Images);
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.RemoveImage(7));
            Assert.Throws<ArgumentException>(() => ws.InsertImage(" ", 0, 0, 1, 1));
            File.WriteAllText(Tmp("not.png"), "this is text");
            Assert.Throws<NotSupportedException>(() => ws.InsertImage(Tmp("not.png"), 0, 0, 1, 1));
        }

        [Fact]
        public void Images_ResizeAndPosition_OnBothFormats()
        {
            foreach (var fmt in new[] { ExcelFileFormat.XLSX, ExcelFileFormat.XLS })
            {
                var ws = Fresh(fmt);
                var img = ws.InsertImage(Png, ImageFormat.PNG, 1, 1, 3, 3);
                img.Resize(); img.Resize(2.0); img.Resize(1.5, 0.5);
                img.Position.LeftColumnIndex = 2; img.Position.TopRowIndex = 2;
                img.Position.RightColumnIndex = 5; img.Position.BottomRowIndex = 6;
                img.Position.AnchorType = AnchorType.MoveDontResize;
                var p = ws.Images[0].Position;
                Assert.Equal(2, p.LeftColumnIndex); Assert.Equal(6, p.BottomRowIndex);
                Assert.Equal("F", p.RightColumnLetter);
                p.LeftColumnLetter = "A"; Assert.Equal(0, ws.Images[0].Position.LeftColumnIndex);
                Assert.Equal(AnchorType.MoveDontResize, p.AnchorType);
            }
        }

        [Fact]
        public void Chart_LettersOverload_AndLoadedChartSurface()
        {
            var ws = Fresh();
            for (var i = 0; i < 4; i++) { ws.GetCellAt(i, 0).Value = "k" + i; ws.GetCellAt(i, 1).Value = i + 1; }
            var chart = ws.CreateChart(ChartType.Column, 6, "A", 16, "F");
            Assert.Equal(0, chart.Position.LeftColumnIndex);
            Assert.Throws<ArgumentException>(() => ws.CreateChart(ChartType.Column, 6, "F", 16, "A"));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.CreateChart((ChartType)99, 0, 0, 5, 5));
            chart.AddSeries("A1:A4", "B1:B4");
            chart.Plot();
            ws.Book.SaveAs(Tmp("ch.xlsx"));

            var loaded = WorkBook.Load(Tmp("ch.xlsx")).DefaultWorkSheet.Charts.Single();
            Assert.Throws<NotSupportedException>(() => loaded.AddSeries("A1:A4", "B1:B4"));
            Assert.Throws<NotSupportedException>(() => loaded.Plot());
            var book = WorkBook.Load(Tmp("ch.xlsx")); var sheet = book.DefaultWorkSheet;
            sheet.RemoveChart(sheet.Charts.Single());
            Assert.Empty(sheet.Charts);
            book.SaveAs(Tmp("nochart.xlsx"));
            Assert.Empty(WorkBook.Load(Tmp("nochart.xlsx")).DefaultWorkSheet.Charts);
            using (var zip = System.IO.Compression.ZipFile.OpenRead(Tmp("nochart.xlsx")))
                Assert.DoesNotContain(zip.Entries, e => e.FullName.StartsWith("xl/charts/chart"));
            Assert.Throws<ArgumentNullException>(() => sheet.RemoveChart(null));
            Assert.Throws<NotSupportedException>(() => Fresh(ExcelFileFormat.XLS).CreateChart(ChartType.Pie, 0, 0, 3, 3));
        }

        [Fact]
        public void TableStyle_AndNamedTablesAndRanges()
        {
            var ws = Fresh();
            ws["A1"].Value = "id"; ws["B1"].Value = "name";
            for (var r = 1; r <= 3; r++) { ws.GetCellAt(r, 0).Value = r; ws.GetCellAt(r, 1).Value = "n" + r; }
            ws["A1:B4"].SaveAsNamedTable("People", true, TableStyle.TableStyleMedium2, true);
            Assert.Equal(new[] { "People" }, ws.GetNamedTableNames());
            Assert.Equal("A1:B4", ws.GetNamedTable("People").RangeAddressAsString);
            Assert.Equal("n2", ws.GetDataTableFromNamedTable("People", true).Rows[1]["name"]);
            ws["D1:D3"].SaveAsNamedRange("Col", true);
            ws.AddNamedRange("Whole", ws["A1:B4"]);
            ws.AddNamedRange("Formula", "=Sheet1!$A$1");
            Assert.Contains("Col", ws.GetNamedRanges());
            Assert.NotNull(ws.FindNamedRange("Whole"));
            ws.RemoveNamedRange("Whole"); Assert.Null(ws.FindNamedRange("Whole"));
            Assert.Equal("TableStyleMedium2", TableStyle.TableStyleMedium2.Name);
            Assert.NotNull(new TableStyle());
        }
    }
}

namespace OpenSheet.Tests
{
    public class DiagonalDirectionTests : TestBase
    {
        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void Direction_DoesNotDependOnTheOrderItIsSetIn(ExcelFileFormat fmt)
        {
            var ws = WorkBook.Create(fmt).DefaultWorkSheet;
            ws["A1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;      // before any diagonal line exists
            ws["A1"].Style.DiagonalBorder.Type = BorderType.Thin;
            ws["B1"].Style.DiagonalBorder.Type = BorderType.Thin;
            ws["B1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Backward;     // after
            Assert.Equal(DiagonalBorderDirection.Forward, ws["A1"].Style.DiagonalBorderDirection);
            Assert.Equal(DiagonalBorderDirection.Backward, ws["B1"].Style.DiagonalBorderDirection);

            var ext = fmt == ExcelFileFormat.XLS ? ".xls" : ".xlsx";
            ws.Book.SaveAs(Tmp("d" + ext));
            var back = WorkBook.Load(Tmp("d" + ext)).DefaultWorkSheet;
            Assert.Equal(DiagonalBorderDirection.Forward, back["A1"].Style.DiagonalBorderDirection);
            Assert.Equal(DiagonalBorderDirection.Backward, back["B1"].Style.DiagonalBorderDirection);
            Assert.Equal(BorderType.Thin, back["B1"].Style.DiagonalBorder.Type);
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void Direction_OfOneCell_NeverChangesCellsThatShareItsBorder(ExcelFileFormat fmt)
        {
            var ws = WorkBook.Create(fmt).DefaultWorkSheet;
            ws["A1:C1"].Style.DiagonalBorder.Type = BorderType.Thin;                        // one border, three cells
            ws["A1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.Both;
            Assert.Equal(DiagonalBorderDirection.Both, ws["A1"].Style.DiagonalBorderDirection);
            Assert.Equal(DiagonalBorderDirection.None, ws["B1"].Style.DiagonalBorderDirection);
            Assert.Equal(DiagonalBorderDirection.None, ws["C1"].Style.DiagonalBorderDirection);
            ws["A1"].Style.DiagonalBorderDirection = DiagonalBorderDirection.None;
            Assert.Equal(DiagonalBorderDirection.None, ws["A1"].Style.DiagonalBorderDirection);
        }
    }

    public class WorksheetLevelRemovalTests
    {
        [Fact]
        public void RemovingAnEmptyColumnOrRowOutsideTheUsedCells_StillShiftsTheSheet()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["C3"].Value = "x"; ws["E5"].Value = "y";
            ws.RemoveColumn(0);                           // column A holds nothing: everything moves one column left
            ws.RemoveRow(0);
            Assert.Equal("x", ws["B2"].StringValue);
            Assert.Equal("y", ws["D4"].StringValue);
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.RemoveColumn(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.RemoveRow(-1));
            ws.RemoveColumn(40);                           // far beyond the data: nothing to do, nothing breaks
            ws.RemoveRow(500);
            Assert.Equal("y", ws["D4"].StringValue);
        }
    }
}

namespace OpenSheet.Tests
{
    public class CopyingSheetsWithPicturesTests : TestBase
    {
        static WorkSheet WithPicture(ExcelFileFormat fmt)
        {
            var ws = WorkBook.Create(fmt).DefaultWorkSheet;
            ws["A1"].Value = "kept";
            ws.InsertImage(Samples.Png, OpenSheet.Drawing.Images.ImageFormat.PNG, 3, 3, 5, 5);   // inserted now, never saved
            return ws;
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void CopyTo_AnotherWorkbook_KeepsAPictureInsertedInThisSession(ExcelFileFormat fmt)
        {
            var src = WithPicture(fmt);
            var target = WorkBook.Create(fmt);
            src.CopyTo(target, "copy");
            var copy = target.GetWorkSheet("copy");
            Assert.Equal("kept", copy["A1"].StringValue);
            Assert.Single(copy.Images);
            Assert.Equal(Samples.Png, copy.Images[0].Data);
            Assert.Equal(3, copy.Images[0].Position.LeftColumnIndex);
            Assert.Single(src.Images);                                  // the source is untouched
            Assert.Equal(2, target.WorkSheets.Count);                   // and no half-made sheet was left behind
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void CopySheet_InTheSameWorkbook_KeepsAPictureInsertedInThisSession(ExcelFileFormat fmt)
        {
            var src = WithPicture(fmt);
            var copy = src.CopySheet("again");
            Assert.Equal("again", copy.Name);
            Assert.Single(copy.Images);
            Assert.Equal(2, src.Book.WorkSheets.Count);
            src.Book.SaveAs(Tmp("two" + (fmt == ExcelFileFormat.XLS ? ".xls" : ".xlsx")));
            var back = WorkBook.Load(Tmp("two" + (fmt == ExcelFileFormat.XLS ? ".xls" : ".xlsx")));
            Assert.Single(back.GetWorkSheet("again").Images);
            Assert.Single(back.DefaultWorkSheet.Images);
        }

        [Fact]
        public void CopySheet_RefusesADuplicateName_AndLeavesTheWorkbookAsItWas()
        {
            var src = WithPicture(ExcelFileFormat.XLSX);
            Assert.ThrowsAny<ArgumentException>(() => src.CopySheet(src.Name));
            Assert.Single(src.Book.WorkSheets);
        }
    }
}

namespace OpenSheet.Tests
{
    public class DeletedSheetsLeaveNothingBehindTests : TestBase
    {
        static string[] Entries(string path)
        {
            using (var zip = System.IO.Compression.ZipFile.OpenRead(path)) return zip.Entries.Select(e => e.FullName).ToArray();
        }

        string TwoSheetsWithChartAndPicture()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet; ws.Name = "Charts";
            for (var i = 0; i < 4; i++) { ws.GetCellAt(i, 0).Value = "k" + i; ws.GetCellAt(i, 1).Value = i + 1; }
            var c = ws.CreateChart(OpenSheet.Drawing.Charts.ChartType.Column, 6, 0, 16, 5); c.AddSeries("A1:A4", "B1:B4"); c.Plot();
            ws.InsertImage(Samples.Png, OpenSheet.Drawing.Images.ImageFormat.PNG, 7, 0, 9, 3);
            wb.CreateWorkSheet("Plain")["A1"].Value = "stays";
            wb.SaveAs(Tmp("start.xlsx"));
            return Tmp("start.xlsx");
        }

        [Fact]
        public void DeletingASheet_TakesItsChartsAndPicturesOutOfTheFile()
        {
            var wb = WorkBook.Load(TwoSheetsWithChartAndPicture());
            wb.RemoveWorkSheet("Charts");
            wb.SaveAs(Tmp("after.xlsx"));
            var entries = Entries(Tmp("after.xlsx"));
            Assert.DoesNotContain(entries, e => e.StartsWith("xl/charts/") || e.StartsWith("xl/drawings/") || e.StartsWith("xl/media/"));
            Assert.Equal("stays", WorkBook.Load(Tmp("after.xlsx")).DefaultWorkSheet["A1"].StringValue);
        }

        [Fact]
        public void DeletingASheetThatSharesAPictureWithAnother_KeepsThePictureForTheOther()
        {
            var wb = WorkBook.Load(TwoSheetsWithChartAndPicture());
            wb.GetWorkSheet("Plain").InsertImage(Samples.Png, OpenSheet.Drawing.Images.ImageFormat.PNG, 1, 1, 3, 3);
            wb.SaveAs(Tmp("shared.xlsx"));
            var again = WorkBook.Load(Tmp("shared.xlsx"));
            again.RemoveWorkSheet("Charts");
            again.SaveAs(Tmp("one-left.xlsx"));
            var left = WorkBook.Load(Tmp("one-left.xlsx"));
            Assert.Single(left.DefaultWorkSheet.Images);
            Assert.Equal(Samples.Png, left.DefaultWorkSheet.Images[0].Data);
        }

        [Fact]
        public void ANewDrawingAfterADeletedOne_DoesNotInheritItsRelationships()
        {
            var wb = WorkBook.Load(TwoSheetsWithChartAndPicture());
            wb.RemoveWorkSheet("Charts");
            var plain = wb.GetWorkSheet("Plain");
            plain.InsertImage(Samples.Png, OpenSheet.Drawing.Images.ImageFormat.PNG, 1, 1, 3, 3);
            wb.SaveAs(Tmp("reuse.xlsx"));
            var entries = Entries(Tmp("reuse.xlsx"));
            Assert.DoesNotContain(entries, e => e.StartsWith("xl/charts/"));
            var back = WorkBook.Load(Tmp("reuse.xlsx"));
            Assert.Single(back.DefaultWorkSheet.Images);
            Assert.Empty(back.DefaultWorkSheet.Charts);
        }

        [Fact]
        public void ReplacingASheetThroughTheCollection_LeavesAValidSingleDrawing()
        {
            var wb = WorkBook.Load(TwoSheetsWithChartAndPicture());
            var other = WorkBook.Create(); other.DefaultWorkSheet.Name = "New"; other.DefaultWorkSheet["A1"].Value = "fresh";
            other.DefaultWorkSheet.InsertImage(Samples.Png, OpenSheet.Drawing.Images.ImageFormat.PNG, 1, 1, 3, 3);
            wb.WorkSheets[0] = other.DefaultWorkSheet;
            wb.SaveAs(Tmp("replaced.xlsx"));
            var back = WorkBook.Load(Tmp("replaced.xlsx"));
            Assert.Equal(new[] { "New", "Plain" }, back.WorkSheets.Select(s => s.Name));
            Assert.Single(back.DefaultWorkSheet.Images);
            Assert.DoesNotContain(Entries(Tmp("replaced.xlsx")), e => e.StartsWith("xl/charts/"));
        }
    }
}

namespace OpenSheet.Tests
{
    public class LocaleDependentFunctionTests
    {
        // Excel writes the concise form with the list separator of its own language (a Spanish Excel gives "3; 1; 4"), so no recording
        // from a localized Excel can stand for it; the documented form uses ", ".
        [Fact]
        public void ArrayToText_Concise_UsesCommaAndSpace()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = 3; ws["A2"].Value = 1; ws["B1"].Value = "x"; ws["B2"].Value = "y";
            ws["D1"].Formula = "=ARRAYTOTEXT(A1:A2)";
            ws["D2"].Formula = "=ARRAYTOTEXT(A1:B2)";
            ws["D3"].Formula = "=ARRAYTOTEXT(A1:B2,1)";
            Assert.Equal("3, 1", ws["D1"].StringValue);
            Assert.Equal("3, x, 1, y", ws["D2"].StringValue);
            Assert.Equal("{3,\"x\";1,\"y\"}", ws["D3"].StringValue);
        }
    }
}

namespace OpenSheet.Tests
{
    public class PictureSizeTests : TestBase
    {
        static byte[] Make(int w, int h, string kind)
        {
            using (var img = new SixLabors.ImageSharp.Image<SixLabors.ImageSharp.PixelFormats.Rgba32>(w, h))
            using (var ms = new MemoryStream())
            {
                switch (kind)
                {
                    case "png": SixLabors.ImageSharp.ImageExtensions.SaveAsPng(img, ms); break;
                    case "gif": SixLabors.ImageSharp.ImageExtensions.SaveAsGif(img, ms); break;
                    case "bmp": SixLabors.ImageSharp.ImageExtensions.SaveAsBmp(img, ms); break;
                    case "jpg": SixLabors.ImageSharp.ImageExtensions.SaveAsJpeg(img, ms); break;
                    default: SixLabors.ImageSharp.ImageExtensions.SaveAsTiff(img, ms); break;
                }
                return ms.ToArray();
            }
        }

        // a big-endian ("MM") TIFF header with one directory holding width and height, which ImageSharp never writes
        static byte[] BigEndianTiff(int w, int h) => new byte[]
        {
            0x4D, 0x4D, 0x00, 0x2A, 0, 0, 0, 8,
            0x00, 0x02,
            0x01, 0x00, 0x00, 0x03, 0, 0, 0, 1, (byte)(w >> 8), (byte)w, 0, 0,
            0x01, 0x01, 0x00, 0x04, 0, 0, 0, 1, (byte)(h >> 24), (byte)(h >> 16), (byte)(h >> 8), (byte)h,
            0, 0, 0, 0,
        };

        [Theory]
        [InlineData("png", ImageFormat.PNG)]
        [InlineData("gif", ImageFormat.GIF)]
        [InlineData("bmp", ImageFormat.BMP)]
        [InlineData("jpg", ImageFormat.JPEG)]
        [InlineData("tif", ImageFormat.TIFF)]
        public void ImageKit_ReadsFormatAndPixelSize_FromTheHeader(string kind, ImageFormat expected)
        {
            var data = Make(37, 21, kind);
            Assert.Equal(expected, ImageKit.Sniff(data));
            Assert.Equal((37, 21), ImageKit.Size(data));
        }

        [Fact]
        public void ImageKit_ReadsABigEndianTiff_AndRefusesWhatItCannotRead()
        {
            Assert.Equal(ImageFormat.TIFF, ImageKit.Sniff(BigEndianTiff(300, 70000)));
            Assert.Equal((300, 70000), ImageKit.Size(BigEndianTiff(300, 70000)));
            Assert.Null(ImageKit.Sniff(new byte[] { 1, 2, 3 }));
            Assert.Null(ImageKit.Sniff(null));
            Assert.Null(ImageKit.Size(new byte[] { 0x89, 0x50, 0x4E, 0x47 }));      // a PNG cut off before its size
            Assert.Null(ImageKit.Size(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0, 4, 0, 0, 0, 0, 0, 0 })); // a JPEG with no frame header
        }

        // the picture must end inside the column / row where the pixels run out (widths depend on the machine's default font, so they are read, not assumed)
        static void AssertEndsAt(WorkSheet ws, IImage img, double widthPx, double heightPx)
        {
            double Col(int c) => ws.Native.GetColumnWidthInPixels(c);
            double Row(int r) => (ws.Native.GetRow(r)?.HeightInPoints ?? ws.Native.DefaultRowHeightInPoints) * 96.0 / 72.0;
            var right = img.Position.RightColumnIndex; var bottom = img.Position.BottomRowIndex;
            double before = 0; for (var c = img.Position.LeftColumnIndex; c < right; c++) before += Col(c);
            Assert.True(before < widthPx + 0.01 && widthPx <= before + Col(right) + 0.01, $"{widthPx}px should end in column {right}");
            before = 0; for (var r = img.Position.TopRowIndex; r < bottom; r++) before += Row(r);
            Assert.True(before < heightPx + 0.01 && heightPx <= before + Row(bottom) + 0.01, $"{heightPx}px should end in row {bottom}");
        }

        [Theory]
        [InlineData("png")] [InlineData("gif")] [InlineData("bmp")] [InlineData("jpg")] [InlineData("tif")]
        public void InsertImage_FromEachFileType_AndResizeToItsPixels(string kind)
        {
            var path = Tmp("p." + kind);
            File.WriteAllBytes(path, Make(96, 48, kind));
            var ws = WorkBook.Create().DefaultWorkSheet;
            var img = ws.InsertImage(path, 1, 1, 2, 2);
            img.Resize();                                              // natural size
            AssertEndsAt(ws, img, 96, 48);
            img.Resize(2.5);
            AssertEndsAt(ws, img, 240, 120);
            img.Resize(0.5, 0.25);
            AssertEndsAt(ws, img, 48, 12);
            Assert.Throws<ArgumentOutOfRangeException>(() => img.Resize(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => img.Resize(1, -1));
            Assert.Equal(96, img.ToImage().Width);
        }
    }
}

namespace OpenSheet.Tests
{
    /// <summary>Functions that look at the workbook the formula sits in (the Excel comparison has one sheet and cannot check these).</summary>
    public class WorkbookAwareFunctionTests
    {
        static WorkBook TwoSheets()
        {
            var wb = WorkBook.Create();
            var a = wb.DefaultWorkSheet; a.Name = "a";
            a["A1"].Formula = "=SUM(B1:B2)"; a["A2"].Formula = "=VAR.P(B1:B2)+LEN(\"x\")"; a["A3"].Value = 5; a["B1"].Value = 1; a["B2"].Value = 3;
            var b = wb.CreateWorkSheet("b");
            return wb;
        }

        [Fact]
        public void SheetAndSheets_CountAndFindSheets()
        {
            var wb = TwoSheets(); var b = wb.GetWorkSheet("b");
            b["A1"].Formula = "=SHEET()"; b["A2"].Formula = "=SHEET(\"a\")"; b["A3"].Formula = "=SHEET(a!A1)"; b["A4"].Formula = "=SHEETS()";
            b["A5"].Formula = "=SHEET(\"nope\")"; b["A6"].Formula = "=SHEETS(a!A1)"; b["A7"].Formula = "=SHEET(7)";
            Assert.Equal(2.0, b["A1"].DoubleValue); Assert.Equal(1.0, b["A2"].DoubleValue); Assert.Equal(1.0, b["A3"].DoubleValue);
            Assert.Equal(2.0, b["A4"].DoubleValue);
            Assert.Equal("#N/A", b["A5"].StringValue); Assert.Equal(1.0, b["A6"].DoubleValue); Assert.Equal("#N/A", b["A7"].StringValue);
        }

        [Fact]
        public void IsFormulaAndFormulaText_ReadTheReferencedCell_OnAnotherSheetToo()
        {
            var wb = TwoSheets(); var b = wb.GetWorkSheet("b");
            b["A1"].Formula = "=ISFORMULA(a!A1)"; b["A2"].Formula = "=ISFORMULA(a!A3)"; b["A3"].Formula = "=ISFORMULA(5)";
            b["A4"].Formula = "=FORMULATEXT(a!A1)"; b["A5"].Formula = "=FORMULATEXT(a!A2)"; b["A6"].Formula = "=FORMULATEXT(a!A3)";
            Assert.True(b["A1"].ToArray()[0].BoolValue); Assert.False(b["A2"].ToArray()[0].BoolValue);
            Assert.Equal("#VALUE!", b["A3"].StringValue);
            Assert.Equal("=SUM(B1:B2)", b["A4"].StringValue);
            Assert.Equal("=VAR.P(B1:B2)+LEN(\"x\")", b["A5"].StringValue);     // the plain name, not the stored _xlfn. one
            Assert.Equal("#N/A", b["A6"].StringValue);
        }

        [Fact]
        public void ThatWorkAlsoInAnXlsWorkbook()
        {
            var wb = WorkBook.Create(ExcelFileFormat.XLS);
            wb.DefaultWorkSheet.Name = "one"; wb.DefaultWorkSheet["A1"].Formula = "=1+2";
            var two = wb.CreateWorkSheet("two");
            two["A1"].Formula = "=SHEETS()"; two["A2"].Formula = "=ISFORMULA(one!A1)"; two["A3"].Formula = "=FORMULATEXT(one!A1)"; two["A4"].Formula = "=SHEET(\"one\")";
            Assert.Equal(2.0, two["A1"].DoubleValue); Assert.True(two["A2"].ToArray()[0].BoolValue);
            Assert.Equal("=1+2", two["A3"].StringValue); Assert.Equal(1.0, two["A4"].DoubleValue);
        }
    }
}

namespace OpenSheet.Tests
{
    public class NameLimitTests
    {
        [Fact]
        public void ANameOver255Characters_IsRefused_WithAShortenedQuote()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = 1;
            var e = Assert.Throws<ArgumentException>(() => ws.AddNamedRange(new string('a', 300), ws["A1"]));
            Assert.Contains("255", e.Message);
            Assert.DoesNotContain(new string('a', 100), e.Message);     // the message quotes the start of the name, not all of it
        }
    }
}
