using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using Xunit.Abstractions;
using CfOp = OpenSheet.Formatting.Enums.ComparisonOperator;
using DvOp = OpenSheet.DataValidations.ComparisonOperator;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Things that hang on cells (merged regions, links, comments, validation, conditional formats, names) must follow the cells
    /// when rows or columns are inserted or removed. The file is built by OpenSheet, edited by OpenSheet, then every feature's
    /// position is compared with what Excel does to the same file.
    /// </summary>
    [Collection("Excel")]
    public class StructureFeaturesOracleTests : TestBase
    {
        readonly ExcelOracle _excel;
        public StructureFeaturesOracleTests(ExcelOracle excel) { _excel = excel; }

        const int Rows = 10, Cols = 14;

        static void Build(WorkBook wb)
        {
            var ws = wb.DefaultWorkSheet; ws.Name = "S";
            for (var i = 0; i < 6; i++) { ws.GetCellAt(i, 0).Value = (double)i; ws.GetCellAt(i, 1).Value = (double)i * 2; }
            ws.Merge("D2:E2");
            ws["D2"].Value = "merged";
            ws.GetCellAt(0, 6).Hyperlink = "https://example.com";
            ws.GetCellAt(2, 6).AddComment("note", "me");
            ws.DataValidations.AddStringListRule("I2:I4", new[] { "a", "b" });
            var rule = ws.ConditionalFormatting.CreateConditionalFormattingRule(CfOp.GreaterThan, "3");
            rule.PatternFormatting.BackgroundColor = "#FFFF00";
            ws.ConditionalFormatting.AddConditionalFormatting("K1:K5", rule);
            ws["K1"].Value = 1.0;
            ws["B1:B4"].SaveAsNamedRange("Pair", true);
        }

        static void BuildExcel(dynamic wb)
        {
            dynamic ws = wb.Worksheets.Item(1); ws.Name = "S";
            for (var i = 0; i < 6; i++) { ws.Cells(i + 1, 1).Value2 = (double)i; ws.Cells(i + 1, 2).Value2 = (double)i * 2; }
            ws.Range("D2:E2").Merge();
            ws.Range("D2").Value2 = "merged";
            ws.Hyperlinks.Add(ws.Range("G1"), "https://example.com");
            ws.Range("G3").AddComment("note");
            ws.Range("I2:I4").Validation.Add(3, 1, 1, "a,b");
            dynamic fc = ws.Range("K1:K5").FormatConditions.Add(1, 5, "=3");
            fc.Interior.Color = 65535;
            ws.Range("K1").Value2 = 1.0;
            wb.Names.Add("Pair", "=S!$B$1:$B$4");
        }

        static string OursAt(WorkSheet ws, int r, int c)
        {
            var cell = ws.GetCellAt(r, c);
            var parts = new List<string>();
            var merged = ws.GetMergedRegions().FirstOrDefault(m => m.RangeAddress.Includes(r, c));
            if (merged != null) parts.Add("merge " + merged.RangeAddressAsString);
            if (cell.Hyperlink != null) parts.Add("link");
            if (cell.Comment != null) parts.Add("comment");
            if (ws.DataValidations.GetRuleForCell(cell) != null) parts.Add("validation");
            var cf = ws.ConditionalFormatting;
            for (var i = 0; i < cf.ConditionalFormattingsCount; i++)
                if (cf.GetConditionalFormattingAt(i).GetFormattingRanges().Any(a => a.Includes(r, c))) { parts.Add("cf"); break; }
            return string.Join(", ", parts);
        }

        static string TheirsAt(dynamic ws, int r, int c)
        {
            dynamic cell = ws.Cells(r + 1, c + 1);
            var parts = new List<string>();
            try { if ((bool)cell.MergeCells) parts.Add("merge " + ((string)cell.MergeArea.Address).Replace("$", "")); } catch (Exception) { }
            try { if ((int)cell.Hyperlinks.Count > 0) parts.Add("link"); } catch (Exception) { }
            try { if (cell.Comment != null) parts.Add("comment"); } catch (Exception) { }
            try { var t = (int)cell.Validation.Type; parts.Add("validation"); } catch (Exception) { }
            try { if ((int)cell.FormatConditions.Count > 0) parts.Add("cf"); } catch (Exception) { }
            return string.Join(", ", parts);
        }

        static readonly ExcelFileFormat[] Formats = { ExcelFileFormat.XLSX, ExcelFileFormat.XLS };

        // Excel treats a sheet the same whichever file it came from, so the one Excel result is the judge of both formats
        void Compare(string scenario, Action<WorkBook> ours, Action<dynamic> theirs)
        {
            foreach (var fmt in Formats) Compare(scenario + " (" + fmt + ")", fmt, ours, theirs);
        }

        void Compare(string scenario, ExcelFileFormat fmt, Action<WorkBook> ours, Action<dynamic> theirs)
        {
            var wb = WorkBook.Create(fmt); Build(wb);
            ours(wb);
            var diffs = new List<string>();
            using (var doc = _excel.New())
            {
                BuildExcel(doc.Book);
                theirs(doc.Book);
                var a = wb.DefaultWorkSheet; dynamic b = doc.Sheet(1);
                for (var r = 0; r < Rows; r++)
                    for (var c = 0; c < Cols; c++)
                    {
                        var mine = OursAt(a, r, c); var excel = TheirsAt(b, r, c);
                        if (mine != excel) diffs.Add($"{a.GetCellAt(r, c).AddressString}: OpenSheet=[{mine}] Excel=[{excel}]");
                    }
                var ourName = a.FindNamedRange("Pair")?.Replace("$", "").Replace("'", "");
                var theirName = ((string)doc.Book.Names.Item("Pair").RefersTo).TrimStart('=').Replace("$", "").Replace("'", "");
                if (!string.Equals(ourName, theirName, StringComparison.OrdinalIgnoreCase)) diffs.Add($"name Pair: OpenSheet=[{ourName}] Excel=[{theirName}]");
            }
            File.AppendAllText(Path.Combine(Path.GetTempPath(), "opensheet-structure-oracle.txt"), $"== {scenario}: {diffs.Count}\n" + string.Join("\n", diffs.Select(d => "   " + d)) + "\n");
            Assert.True(diffs.Count == 0, scenario + ": " + diffs.Count + " difference(s):" + Environment.NewLine + "  " + string.Join(Environment.NewLine + "  ", diffs.Take(25)));
        }

        [ExcelFact] public void InsertRows_MovesEverythingThatHangsOnCells() =>
            Compare("insert 2 rows before row 2", wb => wb.DefaultWorkSheet.InsertRows(1, 2), b => b.Worksheets.Item(1).Rows("2:3").Insert());

        [ExcelFact] public void InsertColumns_MovesEverythingThatHangsOnCells() =>
            Compare("insert 2 columns before D", wb => wb.DefaultWorkSheet.InsertColumns(3, 2), b => b.Worksheets.Item(1).Columns("D:E").Insert());

        [ExcelFact] public void InsertColumn_InsideAMergedRegion_WidensIt() =>
            Compare("insert 1 column before E", wb => wb.DefaultWorkSheet.InsertColumns(4, 1), b => b.Worksheets.Item(1).Columns("E:E").Insert());

        [ExcelFact] public void RemoveColumn_ClosesTheGap_AndFeaturesFollow() =>
            Compare("remove column C", wb => wb.DefaultWorkSheet.RemoveColumn(2), b => b.Worksheets.Item(1).Columns("C:C").Delete());

        [ExcelFact] public void RemoveRow_ClosesTheGap_AndFeaturesFollow() =>
            Compare("remove row 1", wb => wb.DefaultWorkSheet.RemoveRow(0), b => b.Worksheets.Item(1).Rows("1:1").Delete());

        // ---- tables ----

        void CompareTable(string scenario, Action<WorkBook> ours, Action<dynamic> theirs)
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet; ws.Name = "S";
            ws["P1"].Value = "id"; ws["Q1"].Value = "val";
            for (var i = 2; i <= 5; i++) { ws.GetCellAt(i - 1, 15).Value = (double)i; ws.GetCellAt(i - 1, 16).Value = i * 1.5; }
            ws.AddNamedTable("Tbl", ws["P1:Q5"], true, null, true);
            ours(wb);

            string excelRange; string[] excelColumns;
            using (var doc = _excel.New())
            {
                dynamic xs = doc.Sheet(1); xs.Name = "S";
                xs.Range("P1").Value2 = "id"; xs.Range("Q1").Value2 = "val";
                for (var i = 2; i <= 5; i++) { xs.Cells(i, 16).Value2 = (double)i; xs.Cells(i, 17).Value2 = i * 1.5; }
                xs.ListObjects.Add(1, xs.Range("P1:Q5"), Type.Missing, 1).Name = "Tbl";
                theirs(doc.Book);
                dynamic lo = xs.ListObjects.Item("Tbl");
                excelRange = ((string)lo.Range.Address).Replace("$", "");
                // a Spanish Excel names a new table column "Columna1"; the English name is the one OpenSheet writes
                excelColumns = Enumerable.Range(1, (int)lo.ListColumns.Count).Select(i => ((string)lo.ListColumns.Item(i).Name).Replace("Columna", "Column")).ToArray();
            }
            var table = ws.GetNamedTable("Tbl");
            var oursRange = table.RangeAddress.Location;
            var oursColumns = Enumerable.Range(0, table.ColumnCount).Select(i => ws.GetCellAt(table.RangeAddress.FirstRow, table.RangeAddress.FirstColumn + i).StringValue).ToArray();
            File.AppendAllText(Path.Combine(Path.GetTempPath(), "opensheet-structure-oracle.txt"), $"== table: {scenario}: OpenSheet={oursRange} [{string.Join(",", oursColumns)}] Excel={excelRange} [{string.Join(",", excelColumns)}]" + Environment.NewLine);
            Assert.Equal(excelRange, oursRange);
            Assert.Equal(excelColumns, oursColumns);

            // and the file OpenSheet writes is one Excel opens without repair, with the same table
            wb.SaveAs(Tmp("table.xlsx"));
            using (var reopened = _excel.Open(Tmp("table.xlsx")))
            {
                dynamic lo = reopened.Sheet(1).ListObjects.Item("Tbl");
                Assert.Equal(excelRange, ((string)lo.Range.Address).Replace("$", ""));
                Assert.Equal(excelColumns.Length, (int)lo.ListColumns.Count);
            }
        }

        [ExcelFact] public void Table_GrowsWhenRowsAreInsertedInside() =>
            CompareTable("insert 1 row inside", wb => wb.DefaultWorkSheet.InsertRows(3, 1), b => b.Worksheets.Item(1).Rows("4:4").Insert());

        [ExcelFact] public void Table_MovesWhenRowsAreInsertedAbove() =>
            CompareTable("insert 2 rows above", wb => wb.DefaultWorkSheet.InsertRows(0, 2), b => b.Worksheets.Item(1).Rows("1:2").Insert());

        [ExcelFact] public void Table_MovesWhenColumnsAreInsertedBefore() =>
            CompareTable("insert 2 columns before it", wb => wb.DefaultWorkSheet.InsertColumns(3, 2), b => b.Worksheets.Item(1).Columns("D:E").Insert());

        [ExcelFact] public void Table_ShrinksWhenARowInsideIsRemoved() =>
            CompareTable("remove a data row", wb => wb.DefaultWorkSheet.RemoveRow(2), b => b.Worksheets.Item(1).Rows("3:3").Delete());

        [ExcelFact] public void Table_GainsAColumnWhenOneIsInsertedInside() =>
            CompareTable("insert a column inside", wb => wb.DefaultWorkSheet.InsertColumns(16, 1), b => b.Worksheets.Item(1).Columns("Q:Q").Insert());

        [ExcelFact] public void Table_LosesAColumnWhenOneInsideIsRemoved() =>
            CompareTable("remove the last column", wb => wb.DefaultWorkSheet.RemoveColumn(16), b => b.Worksheets.Item(1).Columns("Q:Q").Delete());
    }
}
