using System;
using System.Data;
using System.IO;
using System.Linq;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>
    /// Behaviors the reference library documents, restated as tests (clean room: the wording comes from the published member
    /// descriptions, the expectations are written independently).
    /// </summary>
    public class DocumentedBehaviorTests : TestBase
    {
        static WorkBook New() => WorkBook.Create();

        // sheet["A1"] is a one-cell range; members that exist on a single cell only need the cell itself
        static Cell At(WorkSheet ws, string a1) => ws[a1].ToArray()[0];

        // ---- Cell ----

        [Fact]
        public void Cell_Indexes_AreZeroBased_AndAddressIsExcelNotation()
        {
            var c = New().DefaultWorkSheet.GetCellAt(2, 0);
            Assert.Equal(2, c.RowIndex); Assert.Equal(0, c.ColumnIndex);
            Assert.Equal("A3", c.AddressString); Assert.Equal("A3", c.Location);
        }

        [Fact]
        public void Cell_DateTime_GetsTheDefaultDateTimeFormat_UnlessACustomOneIsGiven()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = new DateTime(2024, 3, 9, 14, 30, 0);
            Assert.Equal("m/d/yy h:mm", ws["A1"].FormatString);
            At(ws, "A2").SetDateTimeValue(new DateTime(2024, 3, 9), "yyyy-mm-dd");
            Assert.Equal("yyyy-mm-dd", ws["A2"].FormatString);
            Assert.Equal("2024-03-09", At(ws, "A2").FormattedCellValue);
            Assert.Equal(new DateTime(2024, 3, 9), ws["A2"].DateTimeValue);
        }

        [Fact]
        public void TimeOfDay_WithNoDate_ReadsAndWritesBackTheSameNumber()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = 0.75; ws["A1"].FormatString = "hh:mm";
            var time = At(ws, "A1").DateTimeValue;
            Assert.Equal(new TimeSpan(18, 0, 0), time.Value.TimeOfDay);
            At(ws, "B1").Value = At(ws, "A1").Value;                 // read a time of day and write it into another cell
            Assert.Equal(0.75, At(ws, "B1").DoubleValue, 9);
            At(ws, "C1").Value = new DateTime(1899, 12, 30, 6, 30, 0);
            Assert.Equal(0.2708333333, At(ws, "C1").DoubleValue, 9);
        }

        [Fact]
        public void ADateBefore1900_IsRefused_NotWrittenAsAnInvalidNumber()
        {
            var cell = At(New().DefaultWorkSheet, "A1");
            Assert.Throws<ArgumentOutOfRangeException>(() => cell.Value = new DateTime(1850, 1, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => cell.Value = DateTime.MinValue);
        }

        [Fact]
        public void Cell_KindFlags()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = 1.5; ws["A2"].Value = "t"; ws["A3"].Value = true; ws["A4"].Value = DateTime.Today; ws["A5"].Formula = "=1/0"; ws["A6"].Formula = "=1+1";
            Assert.True(At(ws, "A1").IsNumeric); Assert.True(At(ws, "A2").IsText); Assert.True(At(ws, "A3").IsBoolean);
            Assert.True(At(ws, "A4").IsDateTime); Assert.True(At(ws, "A5").IsError); Assert.True(At(ws, "A6").IsFormula);
            Assert.True(ws.GetCellAt(10, 10).IsEmpty);
        }

        [Fact]
        public void Cell_NumericGetters_OfNonNumericCell_ReturnZero()
        {
            var c = New().DefaultWorkSheet.GetCellAt(0, 0);
            c.Value = "words";
            Assert.Equal(0, c.IntValue); Assert.Equal(0, c.LongValue); Assert.Equal(0d, c.DoubleValue); Assert.Equal(0m, c.DecimalValue); Assert.Equal(0f, c.FloatValue);
        }

        [Fact]
        public void Cell_Hyperlink_AcceptsTheFourDocumentedKinds_AndCanBeRemoved()
        {
            var ws = New().DefaultWorkSheet;
            foreach (var target in new[] { "https://example.com/a", "ftp://example.com/f", "mailto:a@b.com", "file:///C:/tmp/x.txt", "Sheet1!B2" })
            {
                var c = ws.GetCellAt(0, 0);
                c.Hyperlink = target;
                Assert.Equal(target, c.Hyperlink);
                c.RemoveHyperlink();
                Assert.Null(c.Hyperlink);
            }
            ws.GetCellAt(5, 5).RemoveHyperlink(); // nothing there: no effect
        }

        [Fact]
        public void Cell_Comment_IsNullWhenAbsent_HiddenByDefault_AndRemovingMissingOneDoesNothing()
        {
            var c = New().DefaultWorkSheet.GetCellAt(1, 1);
            Assert.Null(c.Comment);
            c.RemoveComment();
            var cm = c.AddComment("note", "me");
            Assert.False(cm.IsVisible);
            Assert.Equal("B2", cm.Address);
            Assert.Equal("note", c.Comment.Content);
            c.RemoveComment();
            Assert.Null(c.Comment);
        }

        [Fact]
        public void Cell_RemoveFormula_OnPlainCell_DoesNothing()
        {
            var c = New().DefaultWorkSheet.GetCellAt(0, 0);
            c.Value = 4; c.RemoveFormula();
            Assert.Equal(4d, c.Value);
        }

        [Fact]
        public void Cell_Evaluate_StoresTheResult_AndStaysAFormula()
        {
            var wb = New(); var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = 3; ws["B1"].Formula = "=A1*A1";
            At(ws, "B1").Evaluate();
            Assert.True(At(ws, "B1").IsFormula);
            Assert.Equal(9d, ws["B1"].Value);
            At(ws, "A1").Evaluate(); // a plain cell is left alone
            Assert.Equal(3d, ws["A1"].Value);
        }

        [Fact]
        public void Cell_Copy_PlacesValueAndStyleAtThePoint()
        {
            var wb = New(); var ws = wb.DefaultWorkSheet; var other = wb.CreateWorkSheet("Other");
            ws["A1"].Value = "x"; ws["A1"].Style.Font.Bold = true;
            ws["A1"].Copy(other, "C4");
            Assert.Equal("x", other["C4"].Value);
            Assert.True(other["C4"].Style.Font.Bold);
        }

        [Fact]
        public void Cell_GetValue_Throws_WhenTheTypeDoesNotFit_AndTryGetValueDoesNot()
        {
            var c = New().DefaultWorkSheet.GetCellAt(0, 0);
            c.Value = "abc";
            Assert.ThrowsAny<Exception>(() => c.GetValue<int>());
            Assert.False(c.TryGetValue<int>(out var n)); Assert.Equal(0, n);
            c.Value = "42";
            Assert.True(c.TryGetValue<int>(out n)); Assert.Equal(42, n);
        }

        [Fact]
        public void Cell_FormattedCellValue_AppliesTheFormat_NotTheRawNumber()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = 0.256; ws["A1"].FormatString = "0.0%";
            ws["A2"].Value = 1234.5; ws["A2"].FormatString = "#,##0.00";
            ws["A3"].Value = 0.5; // General
            Assert.Equal("25.6%", At(ws, "A1").FormattedCellValue);
            Assert.Equal("1,234.50", At(ws, "A2").FormattedCellValue);
            Assert.Equal("0.5", At(ws, "A3").FormattedCellValue);
            Assert.Equal("0.256", At(ws, "A1").StringValue);   // the cell keeps the raw number, the range shows it formatted
            Assert.Equal("25.6%", ws["A1"].StringValue);
        }

        // ---- Range ----

        [Fact]
        public void Range_Value_OfSeveralCells_IsTheSumOfTheNumbers_AndSetWritesEveryCell()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1:A3"].Value = 2;
            Assert.Equal(6m, ws["A1:A3"].Sum());
            ws["B1"].Value = 1; ws["B2"].Value = "text"; ws["B3"].Value = 4;
            Assert.Equal(5m, Convert.ToDecimal(ws["B1:B3"].Value));
            ws["B1:B3"].Value = "";            // removes the cells entirely
            Assert.True(ws["B1:B3"].IsEmpty);
            Assert.DoesNotContain(ws.FilledCells, c => c.ColumnIndex == 1);
        }

        [Fact]
        public void Range_BoolValue_IsTrueOnlyWhenEveryCellIs_AndPositiveNumbersCount()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = true; ws["A2"].Value = 3; ws["A3"].Value = true;
            Assert.True(ws["A1:A3"].BoolValue);
            ws["A2"].Value = 0;
            Assert.False(ws["A1:A3"].BoolValue);
            ws["A1:A3"].BoolValue = true;
            Assert.True(ws["A1:A3"].BoolValue);
        }

        [Fact]
        public void Range_DateTimeValue_IsTheLatestDate_OrNull()
        {
            var ws = New().DefaultWorkSheet;
            Assert.Null(ws["A1:A3"].DateTimeValue);
            ws["A1"].Value = new DateTime(2020, 1, 1); ws["A2"].Value = new DateTime(2024, 5, 5); ws["A3"].Value = "no date";
            Assert.Equal(new DateTime(2024, 5, 5), ws["A1:A3"].DateTimeValue);
        }

        [Fact]
        public void Range_StringValue_UsesTabsBetweenColumns_AndNewLinesBetweenRows()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = 1; ws["B1"].Value = 2; ws["A2"].Value = 3; ws["B2"].Value = 4;
            Assert.Equal("1\t2" + Environment.NewLine + "3\t4", ws["A1:B2"].StringValue);
        }

        [Fact]
        public void Range_FormatStringAndFormula_AreShared_OrNullWhenTheyDiffer()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1:A2"].FormatString = "0.00";
            Assert.Equal("0.00", ws["A1:A2"].FormatString);
            ws["A2"].FormatString = "0%";
            Assert.Null(ws["A1:A2"].FormatString);
            ws["B1:B2"].Formula = "=1+1";
            Assert.Equal("=1+1", ws["B1:B2"].Formula);
            ws["B2"].Formula = "=2+2";
            Assert.Null(ws["B1:B2"].Formula);
        }

        [Fact]
        public void Range_AddAndSubtractRange_FollowTheDocumentedExamples()
        {
            var ws = New().DefaultWorkSheet;
            Assert.Equal("A1:D4", ws["A1:B2"].AddRange(ws["C3:D4"]).RangeAddressAsString);
            Assert.Equal("A1:D4", (ws["A1:B2"] + ws["C3:D4"]).RangeAddressAsString);

            ws["A1"].Value = "keep"; ws["B2"].Value = "wipe";
            var partial = ws["A1:B2"].SubtractRange(ws["B2:C3"]);   // partial overlap: contents cleared, shape kept
            Assert.Equal("A1:B2", partial.RangeAddressAsString);
            Assert.True(ws["B2"].IsEmpty); Assert.Equal("keep", ws["A1"].Value);

            Assert.Equal("A2:B2", ws["A1:B2"].SubtractRange(ws["A1:C1"]).RangeAddressAsString);  // a full-width slice from the edge
            Assert.True(ws["A1:B2"].SubtractRange(ws["A1:B2"]).RangeAddressAsString == "" || ws["A1:B2"].SubtractRange(ws["A1:B2"]).Count == 0);
        }

        [Fact]
        public void Range_Trim_DropsEmptyRowsAndColumnsOnTheBorders()
        {
            var ws = New().DefaultWorkSheet;
            ws["B2"].Value = 1; ws["C3"].Value = 2;
            Assert.Equal("B2:C3", ws["A1:E5"].Trim().RangeAddressAsString);
        }

        [Fact]
        public void Range_RemoveRow_EmptiesThe_Row_AndMovesRowsUpWhenNothingIsLeft()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = "a"; ws["A2"].Value = "b"; ws["A3"].Value = "c";
            ws.RemoveRow(1);
            Assert.Equal("a", ws["A1"].Value); Assert.Equal("c", ws["A2"].Value); Assert.True(ws["A3"].IsEmpty);
        }

        [Fact]
        public void Range_SetCellValue_UsesRowAndColumnIndexes()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1:C3"].SetCellValue(1, 2, "x");
            Assert.Equal("x", ws["C2"].Value);
        }

        [Fact]
        public void Range_ToDataTable_CanTakeColumnNamesFromTheFirstRow()
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = "id"; ws["B1"].Value = "name"; ws["A2"].Value = 1; ws["B2"].Value = "x";
            var t = ws["A1:B2"].ToDataTable(true);
            Assert.Equal(new[] { "id", "name" }, t.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
            Assert.Single(t.Rows);
            Assert.Equal(2, ws["A1:B2"].ToDataTable(false).Rows.Count);
        }

        [Fact]
        public void RangeAddress_Operations()
        {
            var ws = New().DefaultWorkSheet;
            var a = ws["B2:D4"].RangeAddress; var b = ws["C3:F6"].RangeAddress;
            Assert.Equal("C3:D4", a.GetIntersection(b).Location);
            Assert.Null(a.GetIntersection(ws["H1:H2"].RangeAddress));
            Assert.Equal("B2:F6", a.Extend(b).Location);
            Assert.True(a.Includes(3, 3)); Assert.False(a.Includes(0, 0));
            Assert.True(a.IsEqual(ws["B2:D4"].RangeAddress));
            Assert.Equal(3, a.RowsCount); Assert.Equal(3, a.ColumnsCount);
        }

        // ---- WorkSheet ----

        [Fact]
        public void PhysicalRowCount_CountsRowsHoldingCellsOrOnlyFormatting()
        {
            var ws = New().DefaultWorkSheet;
            ws["A11"].Value = 1;
            ws.GetRow(19).Hidden = true; // row 20 has no cell but carries a setting
            Assert.Equal("A11:A20", ws.RangeAddress.Location);
            Assert.Equal(2, ws.PhysicalRowCount);
        }

        [Fact]
        public void PhysicalColumnCount_CountsColumnsHoldingCellsOrOnlyFormatting()
        {
            var ws = New().DefaultWorkSheet;
            ws["B1"].Value = 1;
            ws.GetColumn("K").Hidden = true; // column K has no cell but carries a setting
            Assert.Equal("B1", ws.RangeAddress.Location.Split(':')[0]);
            Assert.Equal(2, ws.PhysicalColumnCount);
        }

        [Fact]
        public void FirstAndLastFilledCell_AreNullOnAnEmptySheet()
        {
            var ws = New().DefaultWorkSheet;
            Assert.Null(ws.FirstFilledCell); Assert.Null(ws.LastFilledCell); Assert.Empty(ws.FilledCells);
            ws["C3"].Value = 1; ws["B5"].Value = 2;
            Assert.Equal("C3", ws.FirstFilledCell.AddressString); Assert.Equal("B5", ws.LastFilledCell.AddressString);
        }

        [Fact]
        public void WholeColumnAndRowAddresses_CoverTheCellsInUse()
        {
            var ws = New().DefaultWorkSheet;
            ws["B2"].Value = 1; ws["D5"].Value = 2;
            Assert.Equal("A2:A5", ws["A:A"].RangeAddressAsString);
            Assert.Equal("B2:C5", ws["B:C"].RangeAddressAsString);
            Assert.Equal("B3:D3", ws["3:3"].RangeAddressAsString);
            Assert.Equal("B2:D4", ws["2:4"].RangeAddressAsString);
            Assert.Equal(3m, ws["B:D"].Sum());
        }

        [Fact]
        public void TabColor_IsNullUntilSet()
        {
            var ws = New().DefaultWorkSheet;
            Assert.Null(ws.TabColor);
            ws.SetWorksheetTabColor(new byte[] { 1, 2, 3 });
            Assert.Equal(new byte[] { 1, 2, 3 }, ws.TabColor);
        }

        [Fact]
        public void Merge_ReturnsAnIndex_AndUnmergeWorksByIndexOrByRange()
        {
            var ws = New().DefaultWorkSheet;
            var i = ws.Merge("A1:B1"); ws.Merge("C1:D1");
            Assert.Equal(2, ws.GetMergedRegions().Count);
            ws.Unmerge(i);
            Assert.Single(ws.GetMergedRegions());
            ws.Unmerge("C1:D1");
            Assert.Empty(ws.GetMergedRegions());
            ws.Unmerge("E1:F1"); // not merged: no effect
        }

        [Theory]
        [InlineData("T1")] [InlineData("A1")] [InlineData("XFD1048576")] [InlineData("R1C1")] [InlineData("R")] [InlineData("c")] [InlineData("my name")]
        [InlineData("1abc")] [InlineData("a-b")] [InlineData("")]
        public void NamedRangesAndTables_RefuseNamesExcelWouldRejectAsDamage(string name)
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = "h"; ws["A2"].Value = 1;
            Assert.Throws<ArgumentException>(() => ws.AddNamedTable(name, ws["A1:A2"], true, null, true));
            Assert.Throws<ArgumentException>(() => ws.AddNamedRange(name, ws["A1:A2"]));
        }

        [Theory]
        [InlineData("Sales")] [InlineData("_hidden")] [InlineData("Total.2024")] [InlineData("XFE1")] [InlineData("AAAA1")] [InlineData("Ñandú")]
        public void NamedRangesAndTables_AcceptOrdinaryNames(string name)
        {
            var ws = New().DefaultWorkSheet;
            ws["A1"].Value = "h"; ws["A2"].Value = 1;
            ws.AddNamedRange(name, ws["A1:A2"]);
            Assert.Contains(name, ws.GetNamedRanges());
            ws.AddNamedTable("T" + name, ws["A1:A2"], true, null, true);
        }

        // ---- WorkBook ----

        [Fact]
        public void EvaluateAll_ComputesEveryFormula_AndKeepsThemFormulas()
        {
            var wb = New(); var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = 2; ws["A2"].Formula = "=A1+1"; ws["A3"].Formula = "=A2*2";
            Assert.Same(wb, wb.EvaluateAll());
            Assert.True(At(ws, "A3").IsFormula); Assert.Equal(6d, ws["A3"].Value);
        }

        [Fact]
        public void AccellerateFormulaEvaluation_KeepsResultsUntilEvaluateAllIsCalled()
        {
            var wb = New(); var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = 1; ws["B1"].Formula = "=A1*2";
            Assert.Equal(2d, ws["B1"].Value);
            ws["A1"].Value = 5;
            Assert.Equal(10d, ws["B1"].Value);          // by default every edit refreshes
            wb.AccellerateFormulaEvaluation = true;
            Assert.Equal(10d, ws["B1"].Value);
            ws["A1"].Value = 7;
            Assert.Equal(10d, ws["B1"].Value);          // cached on purpose
            wb.EvaluateAll();
            Assert.Equal(14d, ws["B1"].Value);
        }

        [Fact]
        public void SaveAsWithCustomDelimiter_WritesThatSeparator()
        {
            var wb = New();
            wb.DefaultWorkSheet["A1"].Value = "a"; wb.DefaultWorkSheet["B1"].Value = "b";
            wb.SaveAsWithCustomDelimiter(Tmp("x.csv"), ";");
            Assert.Equal("a;b", File.ReadAllText(Tmp("x.csv")).Trim());
        }

        [Fact]
        public void DataSet_RoundTrip_OneSheetPerTable()
        {
            var wb = New(); var ws = wb.DefaultWorkSheet; ws.Name = "People";
            ws["A1"].Value = "id"; ws["A2"].Value = 1; ws["A3"].Value = 2;
            var other = wb.CreateWorkSheet("Empty");
            var ds = wb.ToDataSet(true);
            Assert.Equal(2, ds.Tables.Count);
            Assert.Equal("People", ds.Tables[0].TableName);
            Assert.Equal(2, ds.Tables[0].Rows.Count);

            var back = WorkBook.Load(ds);
            Assert.Equal("People", back.DefaultWorkSheet.Name);
            var added = back.LoadWorkSheet(ds.Tables[0]);
            Assert.NotNull(back.GetWorkSheet(added.Name));
        }

        [Fact]
        public void SetSheetPosition_AndActiveTab()
        {
            var wb = New();
            wb.CreateWorkSheet("B"); wb.CreateWorkSheet("C");
            wb.SetSheetPosition("C", 0);
            Assert.Equal("C", wb.WorkSheets[0].Name);
            wb.SetActiveTab(1);
            Assert.Equal(3, wb.WorkSheets.Count);
        }

        [Fact]
        public void GetWorkSheet_ReturnsNullForAnUnknownName()
        {
            Assert.Null(New().GetWorkSheet("nope"));
        }
    }
}
