using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;
using Xunit.Abstractions;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// The same edit done by OpenSheet and by Excel on the same data; every cell of the result (formula text or value) is compared.
    /// Formulas make this matter: references must follow the cells they point at, as they do in Excel.
    /// </summary>
    [Collection("Excel")]
    public class OperationOracleTests : TestBase
    {
        readonly ExcelOracle _excel;
        readonly ITestOutputHelper _out;
        public OperationOracleTests(ExcelOracle excel, ITestOutputHelper output) { _excel = excel; _out = output; }

        static readonly object ReportLock = new object();

        const int Rows = 12, Cols = 9;

        // A1:A5 numbers, B = twice A, C1 total, C2 absolute+relative mix, D3 pointer into the middle, E1 range sum with a text column, second sheet reads the first.
        static void Build(WorkBook wb)
        {
            var ws = wb.DefaultWorkSheet; ws.Name = "S";
            var other = wb.CreateWorkSheet("T");
            for (var i = 0; i < 5; i++) { ws.GetCellAt(i, 0).Value = (double)(i + 1) * 10; ws.GetCellAt(i, 1).Formula = $"=A{i + 1}*2"; }
            ws["C1"].Formula = "=SUM(A1:A5)";
            ws["C2"].Formula = "=$A$1+B2";
            ws["D3"].Formula = "=A3+B4";
            ws["E1"].Formula = "=SUM(A2:B4)";
            ws["F1"].Formula = "=MAX(A1:A5)-MIN(B1:B5)";
            ws["G1"].Formula = "=T!A1+S!A2";
            other["A1"].Value = 7.0; other["B1"].Formula = "=S!A3+S!B3"; other["C1"].Formula = "=SUM(S!A1:A5)";
        }

        static void BuildExcel(dynamic book)
        {
            dynamic ws = book.Worksheets.Item(1); ws.Name = "S";
            dynamic other = book.Worksheets.Add(Type.Missing, ws); other.Name = "T";
            for (var i = 0; i < 5; i++) { ws.Cells(i + 1, 1).Value2 = (double)(i + 1) * 10; ws.Cells(i + 1, 2).Formula = $"=A{i + 1}*2"; }
            ws.Range("C1").Formula = "=SUM(A1:A5)";
            ws.Range("C2").Formula = "=$A$1+B2";
            ws.Range("D3").Formula = "=A3+B4";
            ws.Range("E1").Formula = "=SUM(A2:B4)";
            ws.Range("F1").Formula = "=MAX(A1:A5)-MIN(B1:B5)";
            ws.Range("G1").Formula = "=T!A1+S!A2";
            other.Range("A1").Value2 = 7.0; other.Range("B1").Formula = "=S!A3+S!B3"; other.Range("C1").Formula = "=SUM(S!A1:A5)";
        }

        // The old format stores a deleted reference without its sheet, so "S!#REF!" and "#REF!" are the same thing there.
        static string Norm(string formula) => formula == null ? null : System.Text.RegularExpressions.Regex.Replace(formula.Replace(" ", "").Replace("'", "").ToUpperInvariant(), @"[A-Z0-9_]+!#REF!", "#REF!");

        static string Ours(WorkSheet ws, int r, int c)
        {
            var cell = ws.GetCellAt(r, c);
            if (cell.IsFormula) return Norm(cell.Formula) + "  = " + Show(cell.Value);
            return cell.IsEmpty ? "" : Show(cell.Value);
        }

        static string Theirs(dynamic ws, int r, int c)
        {
            dynamic cell = ws.Cells(r + 1, c + 1);
            if ((bool)cell.HasFormula)
            {
                object v = cell.Value2;
                return Norm((string)cell.Formula) + "  = " + Show(v is int e && e < -2000000000 ? (object)Diffs.ExcelErrorText(e) : v);
            }
            object val = cell.Value2;
            return val == null ? "" : Show(val is int ev && ev < -2000000000 ? (object)Diffs.ExcelErrorText(ev) : val);
        }

        static string Show(object v)
        {
            if (v == null) return "";
            if (v is DateTime dt) v = NPOI.SS.UserModel.DateUtil.GetExcelDate(dt);
            if (v is double d) return Math.Round(d, 9).ToString("G10", CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        // every scenario runs on both file formats; the expectation is always what Excel does
        void Compare(string scenario, Action<WorkBook> ours, Action<dynamic> theirs, params string[] knownDifferentCells)
        {
            foreach (var format in new[] { ExcelFileFormat.XLSX, ExcelFileFormat.XLS })
                CompareIn(format, scenario + " [" + format + "]", ours, theirs, knownDifferentCells);
        }

        void CompareIn(ExcelFileFormat format, string scenario, Action<WorkBook> ours, Action<dynamic> theirs, params string[] knownDifferentCells)
        {
            var wb = WorkBook.Create(format); Build(wb);
            ours(wb);

            var report = new StringBuilder();
            var diffs = new List<string>();
            using (var doc = _excel.New())
            {
                BuildExcel(doc.Book);
                theirs(doc.Book);
                foreach (var sheetName in new[] { "S", "T" })
                {
                    var a = wb.GetWorkSheet(sheetName); dynamic b = doc.Sheet(sheetName);
                    for (var r = 0; r < Rows; r++)
                        for (var c = 0; c < Cols; c++)
                        {
                            string mine, excel;
                            try { mine = Ours(a, r, c); } catch (Exception e) { mine = "THROWS " + e.GetType().Name; }
                            try { excel = Theirs(b, r, c); } catch (Exception e) { excel = "EXCEL " + e.Message.Split('\n')[0]; }
                            var address = sheetName + "!" + a.GetCellAt(r, c).AddressString;
                            if (mine == excel) continue;
                            if (knownDifferentCells.Contains(address)) continue;
                            diffs.Add($"{address}: OpenSheet=[{mine}] Excel=[{excel}]");
                        }
                }
            }
            File.AppendAllText(Path.Combine(Path.GetTempPath(), "opensheet-operation-oracle.txt"), $"== {scenario}: {diffs.Count} difference(s)\n" + string.Join("\n", diffs.Select(d => "   " + d)) + "\n");
            Assert.True(diffs.Count == 0, scenario + ": " + diffs.Count + " difference(s) from Excel:" + Environment.NewLine + "  " + string.Join(Environment.NewLine + "  ", diffs.Take(25)));
        }

        [ExcelFact] public void InsertRow_MovesCellsAndFollowsReferences() =>
            Compare("insert 1 row before row 3", wb => wb.DefaultWorkSheet.InsertRows(2, 1), b => b.Worksheets.Item("S").Rows("3:3").Insert());

        [ExcelFact] public void InsertTwoRows_AtTheTop() =>
            Compare("insert 2 rows at the top", wb => wb.DefaultWorkSheet.InsertRows(0, 2), b => b.Worksheets.Item("S").Rows("1:2").Insert());

        [ExcelFact] public void InsertColumns_MovesCellsAndFollowsReferences() =>
            Compare("insert 2 columns before B", wb => wb.DefaultWorkSheet.InsertColumns(1, 2), b => b.Worksheets.Item("S").Columns("B:C").Insert());

        [ExcelFact] public void RemoveRow_ClosesTheGap_AndReferencesToItBecomeErrors() =>
            Compare("remove row 3", wb => wb.DefaultWorkSheet.RemoveRow(2), b => b.Worksheets.Item("S").Rows("3:3").Delete());

        [ExcelFact] public void RemoveColumn_ClosesTheGap() =>
            Compare("remove column B", wb => wb.DefaultWorkSheet.RemoveColumn(1), b => b.Worksheets.Item("S").Columns("B:B").Delete());

        [ExcelFact] public void CopyRange_ShiftsRelativeReferences() =>
            Compare("copy B1:B3 to E5", wb => wb.DefaultWorkSheet["B1:B3"].Copy(wb.DefaultWorkSheet, "E5"), b => b.Worksheets.Item("S").Range("B1:B3").Copy(b.Worksheets.Item("S").Range("E5")));

        [ExcelFact] public void CopyRange_ToAnotherSheet() =>
            Compare("copy B1:C2 to T!E5", wb => wb.DefaultWorkSheet["B1:C2"].Copy(wb.GetWorkSheet("T"), "E5"), b => b.Worksheets.Item("S").Range("B1:C2").Copy(b.Worksheets.Item("T").Range("E5")));

        [ExcelFact] public void CopyCell_ShiftsRelativeReferences() =>
            Compare("copy D3 to F8", wb => wb.DefaultWorkSheet["D3"].Copy(wb.DefaultWorkSheet, "F8"), b => b.Worksheets.Item("S").Range("D3").Copy(b.Worksheets.Item("S").Range("F8")));

        [ExcelFact] public void SortDescending_KeepsEachRowsFormulaOnItsRow() =>
            Compare("sort A1:B5 descending by A", wb => wb.DefaultWorkSheet["A1:B5"].SortByColumn(0, SortOrder.Descending),
                b => b.Worksheets.Item("S").Range("A1:B5").Sort(b.Worksheets.Item("S").Range("A1"), 2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2));
    }
}
