using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Xunit;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Files written by Excel itself (shared strings, styles Excel chose, cached formula results) are read by OpenSheet's loader and by
    /// its streaming reader; every cell must come out as Excel holds it.
    /// </summary>
    [Collection("Excel")]
    public class ExcelFileReadingOracleTests : TestBase
    {
        readonly ExcelOracle _excel;
        public ExcelFileReadingOracleTests(ExcelOracle excel) { _excel = excel; }

        const int Rows = 14, Cols = 6;

        void Build(string path, int fileFormat = 51, string password = null)
        {
            using (var doc = _excel.New())
            {
                dynamic ws = doc.Sheet(1); ws.Name = "Mixed";
                ws.Range("A1").Value2 = 42.5;
                ws.Range("B1").Value2 = "plain";
                ws.Range("C1").Value2 = true;
                ws.Range("D1").Value2 = "  leading & trailing  ";
                ws.Range("E1").Value2 = "line1\nline2 <tag> \"quoted\"";
                ws.Range("F1").Value2 = "ñandú ✓ 日本語";
                ws.Range("A2").Value2 = 45000.0; ws.Range("A2").NumberFormat = "yyyy-mm-dd";
                ws.Range("B2").Value2 = 0.75; ws.Range("B2").NumberFormat = "hh:mm";
                ws.Range("C2").Value2 = 45000.5; ws.Range("C2").NumberFormat = "yyyy-mm-dd hh:mm";
                ws.Range("D2").Value2 = 1234567.891; ws.Range("D2").NumberFormat = "#,##0.00";
                ws.Range("E2").Value2 = 0.256; ws.Range("E2").NumberFormat = "0.0%";
                ws.Range("F2").Value2 = 1e-12;
                ws.Range("A3").Formula = "=A1*2";
                ws.Range("B3").Formula = "=1/0";
                ws.Range("C3").Formula = "=\"a\"&\"b\"";
                ws.Range("D3").Formula = "=A2+1"; ws.Range("D3").NumberFormat = "yyyy-mm-dd";
                ws.Range("E3").Formula = "=NA()";
                ws.Range("F3").Formula = "=A1>1";
                // gaps: row 4 empty, row 5 only column D, row 6 only column A and F
                ws.Range("D5").Value2 = "gap";
                ws.Range("A6").Value2 = 1.0; ws.Range("F6").Value2 = 2.0;
                ws.Range("A7").Value2 = "123"; ws.Range("A7").NumberFormat = "@";   // a number written as text
                ws.Range("B7").Value2 = 9007199254740993.0;
                ws.Range("C7").Value2 = -0.0;
                ws.Range("A8").Value2 = ""; // an empty string is an empty cell
                ws.Range("A9").Value2 = new DateTime(1900, 3, 1);
                ws.Range("B9").Value2 = new DateTime(2024, 2, 29);
                if (password == null) doc.Book.SaveAs(path, fileFormat);
                else doc.Book.SaveAs(path, fileFormat, password);
            }
        }

        static string Expected(dynamic cell)
        {
            object v = cell.Value2;
            if (v == null) return "empty";
            if (v is int e && e < -2000000000) return "error";
            if (v is bool b) return "bool:" + b;
            if (v is double d) return "num:" + d.ToString("R", CultureInfo.InvariantCulture);
            return "str:" + (string)v;
        }

        static string FromObject(object value, string kind)
        {
            switch (value)
            {
                case null: return "empty";
                case DateTime dt: return "num:" + Internal.Serial.FromDateTime(dt).ToString("R", CultureInfo.InvariantCulture);
                case double d: return "num:" + d.ToString("R", CultureInfo.InvariantCulture);
                case bool b: return "bool:" + b;
                case string s when kind == "error" || (s.StartsWith("#") && s.Length > 2 && (s.EndsWith("!") || s.EndsWith("?") || s.EndsWith("A"))): return "error";
                default: return "str:" + value;
            }
        }

        [ExcelFact]
        public void Loader_ReadsEveryCellTheWayExcelHoldsIt()
        {
            Build(Tmp("excel.xlsx"));
            var book = WorkBook.Load(Tmp("excel.xlsx"));
            var ws = book.DefaultWorkSheet;
            var diffs = new List<string>();
            using (var doc = _excel.Open(Tmp("excel.xlsx")))
            {
                dynamic xs = doc.Sheet(1);
                for (var r = 0; r < Rows; r++)
                    for (var c = 0; c < Cols; c++)
                    {
                        var expected = Expected(xs.Cells(r + 1, c + 1));
                        var cell = ws.GetCellAt(r, c);
                        var ours = cell.IsError ? "error" : FromObject(cell.Value, "");
                        if (expected == "str:" && ours == "empty") continue; // an empty string is an empty cell
                        if (expected != ours) diffs.Add($"{cell.AddressString}: Excel=[{expected}] OpenSheet=[{ours}]");
                    }
            }
            Assert.True(diffs.Count == 0, diffs.Count + " cell(s) read differently:" + Environment.NewLine + "  " + string.Join(Environment.NewLine + "  ", diffs));
        }

        [ExcelFact]
        public void StreamingReader_ReadsEveryCellTheWayExcelHoldsIt()
        {
            Build(Tmp("excel.xlsx"));
            var streamed = new Dictionary<(int, int), StreamCell>();
            foreach (var row in WorkBook.StreamRows(Tmp("excel.xlsx")))
                foreach (var cell in row.Cells) streamed[(row.RowNumber, cell.ColumnIndex)] = cell;

            var diffs = new List<string>();
            using (var doc = _excel.Open(Tmp("excel.xlsx")))
            {
                dynamic xs = doc.Sheet(1);
                for (var r = 0; r < Rows; r++)
                    for (var c = 0; c < Cols; c++)
                    {
                        var expected = Expected(xs.Cells(r + 1, c + 1));
                        string ours;
                        if (!streamed.TryGetValue((r, c), out var sc) || sc.Type == StreamCellType.Blank) ours = "empty";
                        else if (sc.Type == StreamCellType.Error) ours = "error";
                        else ours = FromObject(sc.Value, "");
                        if (expected == "str:" && ours == "empty") continue;
                        if (expected != ours) diffs.Add($"R{r + 1}C{c + 1}: Excel=[{expected}] streaming=[{ours}]");
                    }
            }
            Assert.True(diffs.Count == 0, diffs.Count + " cell(s) streamed differently:" + Environment.NewLine + "  " + string.Join(Environment.NewLine + "  ", diffs));
        }

        [ExcelFact]
        public void StyleAndDisplay_OfAnExcelWrittenFile_MatchWhatExcelDraws()
        {
            Build(Tmp("excel.xlsx"));
            var ws = WorkBook.Load(Tmp("excel.xlsx")).DefaultWorkSheet;
            var diffs = new List<string>();
            using (var doc = _excel.Open(Tmp("excel.xlsx")))
            {
                dynamic xs = doc.Sheet(1);
                xs.Columns("A:F").ColumnWidth = 40;
                foreach (var a in new[] { "A1", "A2", "B2", "C2", "D2", "E2", "F2", "D3", "A9", "B9" }) // not C1: a boolean is drawn in the language of Excel
                {
                    var ours = ws[a].ToArray()[0].FormattedCellValue;
                    var theirs = (string)xs.Range(a).Text;
                    if (ours != theirs) diffs.Add($"{a}: Excel draws [{theirs}] OpenSheet [{ours}]");
                }
            }
            Assert.True(diffs.Count == 0, string.Join(Environment.NewLine, diffs));
        }

        [ExcelFact]
        public void LoaderReadsAnXlsWrittenByExcel()
        {
            Build(Tmp("excel.xls"), 56);
            var ws = WorkBook.Load(Tmp("excel.xls")).DefaultWorkSheet;
            var diffs = new List<string>();
            using (var doc = _excel.Open(Tmp("excel.xls")))
            {
                dynamic xs = doc.Sheet(1);
                for (var r = 0; r < Rows; r++)
                    for (var c = 0; c < Cols; c++)
                    {
                        var expected = Expected(xs.Cells(r + 1, c + 1));
                        var cell = ws.GetCellAt(r, c);
                        var ours = cell.IsError ? "error" : FromObject(cell.Value, "");
                        if (expected == "str:" && ours == "empty") continue;
                        if (expected != ours) diffs.Add($"{cell.AddressString}: Excel=[{expected}] OpenSheet=[{ours}]");
                    }
            }
            Assert.True(diffs.Count == 0, diffs.Count + " cell(s) read differently:" + Environment.NewLine + "  " + string.Join(Environment.NewLine + "  ", diffs));
        }

        [ExcelFact]
        public void LoaderOpensAFileEncryptedByExcel_WithItsPassword_AndRefusesOthers()
        {
            Build(Tmp("locked.xlsx"), 51, "s3cret!");
            var ws = WorkBook.Load(Tmp("locked.xlsx"), "s3cret!").DefaultWorkSheet;
            Assert.Equal(42.5, ws["A1"].ToArray()[0].DoubleValue);
            Assert.Equal("plain", ws["B1"].ToArray()[0].StringValue);
            Assert.ThrowsAny<Exception>(() => WorkBook.Load(Tmp("locked.xlsx"), "wrong"));
            Assert.ThrowsAny<Exception>(() => WorkBook.Load(Tmp("locked.xlsx")));
        }
    }
}
