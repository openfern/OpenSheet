using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OpenSheet.Internal;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>The shortcuts taken for speed, each held to the slower route it stands in for.</summary>
    public class FastPathTests : TestBase
    {
        // ---- General format ----

        [Fact]
        public void General_WholeNumbers_AgreeWithTheFullAlgorithm()
        {
            var rnd = new Random(7);
            var samples = new double[] { 0, -0.0, 1, -1, 9, 10, 99, 100, 12345, 99999999999, -99999999999, 10000000000, 100000000000, 123456789012, 1e11 - 1, 1e11, 2.5, -0.5, 1e-7, 1e15 };
            foreach (var v in samples) Assert.Equal(Display.GeneralFull(v), Display.General(v));
            for (var i = 0; i < 20000; i++)
            {
                var digits = rnd.Next(1, 13);
                var v = Math.Floor(rnd.NextDouble() * Math.Pow(10, digits)) * (rnd.Next(2) == 0 ? 1 : -1);
                Assert.Equal(Display.GeneralFull(v), Display.General(v));
            }
        }

        // ---- cell text ----

        [Fact]
        public void FormattedCellValue_OfTextBooleansAndBlanks_MatchesNpoi()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "  spaced  ";
            ws["A2"].Value = "é☃ multi\nline";
            ws["A3"].Value = true;
            ws["A4"].Value = false;
            ws["A5"].Value = "x";
            ws["A5"].ClearContents(); // a cell that exists but is blank
            ws["A6"].Value = 12;
            ws["A7"].Formula = "=A6&\"!\"";
            var formatter = new NPOI.SS.UserModel.DataFormatter(CultureInfo.InvariantCulture);
            for (var r = 0; r < 7; r++)
            {
                var native = ws.Native.GetRow(r).GetCell(0);
                Assert.Equal(formatter.FormatCellValue(native, wb.Evaluator), ws.GetCellAt(r, 0).FormattedCellValue);
            }
            Assert.Equal("", ws.GetCellAt(50, 50).FormattedCellValue);
        }

        [Fact]
        public void DateFormattedNumbers_NoDateCanHold_ReadAsNumbers_InsteadOfThrowing()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var serials = new[] { 2958465.5, -1d, 2958466d, 0.25 };
            for (var i = 0; i < serials.Length; i++)
            {
                ws.GetCellAt(i, 0).Value = serials[i];
                ws.GetCellAt(i, 0).FormatString = "yyyy-mm-dd";
            }
            Assert.IsType<double>(ws.GetCellAt(0, 0).Value);
            Assert.IsType<double>(ws.GetCellAt(1, 0).Value);
            Assert.IsType<double>(ws.GetCellAt(2, 0).Value);
            Assert.IsType<DateTime>(ws.GetCellAt(3, 0).Value); // a time of day is still a date-time
            Assert.Equal("########", ws.GetCellAt(1, 0).FormattedCellValue);
            Assert.Equal(serials[0] + serials[1] + serials[2], (double)ws["A1:A3"].Sum());
        }

        // ---- date format codes ----

        [Theory]
        [InlineData(45292.5, "yyyy-mm-dd hh:mm:ss", "2024-01-01 12:00:00")]
        [InlineData(45292.5, "h:mm:ss AM/PM", "12:00:00 PM")]
        [InlineData(45292.25, "h:mm a/p", "6:00 a")]
        [InlineData(45292, "yy/m/d", "24/1/1")]
        [InlineData(45292, "dddd, mmmm d", "Monday, January 1")]
        [InlineData(1.5, "[h]:mm", "36:00")]
        [InlineData(1.5, "[hh]:mm:ss", "36:00:00")]
        [InlineData(0.123456, "s.000", "46.598")]
        [InlineData(0.5, "hh:mm;@", "12:00")]
        [InlineData(45292, "\"Day \"d", "Day 1")]
        public void DateFormat_Draws(double serial, string code, string expected)
        {
            Assert.True(DateFormat.TryFormat(serial, code, out var text));
            Assert.Equal(expected, text);
            Assert.True(DateFormat.TryFormat(serial, code, out var again)); // second time comes from the cache
            Assert.Equal(expected, again);
        }

        [Theory]
        [InlineData("0.00")]
        [InlineData("#,##0")]
        [InlineData("General")]
        [InlineData("@")]
        public void DateFormat_RefusesNumberCodes_EveryTime(string code)
        {
            Assert.False(DateFormat.TryFormat(45292, code, out _));
            Assert.False(DateFormat.TryFormat(45292, code, out _));
        }

        [Fact]
        public void DateFormat_ManyDistinctCodes_StayCorrectPastTheCacheLimit()
        {
            for (var i = 0; i < 1500; i++)
            {
                Assert.True(DateFormat.TryFormat(45292, "yyyy\"-c" + i + "-\"mm", out var text));
                Assert.Equal("2024-c" + i + "-01", text);
            }
        }

        [Fact]
        public void DateFormat_IsSafeAcrossThreads()
        {
            var failures = 0;
            Parallel.For(0, 8, t =>
            {
                for (var i = 0; i < 400; i++)
                {
                    var code = "yyyy\"-t" + t + "_" + (i % 50) + "-\"dd";
                    if (!DateFormat.TryFormat(45292, code, out var text) || text != "2024-t" + t + "_" + (i % 50) + "-01") System.Threading.Interlocked.Increment(ref failures);
                }
            });
            Assert.Equal(0, failures);
        }

        // ---- range format ----

        [Fact]
        public void RangeFormatString_StylesEveryCell_WithOneSharedStyle()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            for (var r = 0; r < 20; r++) ws.GetCellAt(r, 0).Value = r + 0.5;
            var before = wb.Native.NumCellStyles;
            ws["A1:A20"].FormatString = "0.000";
            Assert.Equal(before + 1, wb.Native.NumCellStyles);
            Assert.All(Enumerable.Range(0, 20), r => Assert.Equal("0.000", ws.GetCellAt(r, 0).FormatString));
            Assert.Equal("0.500", ws.GetCellAt(0, 0).FormattedCellValue);
            ws["A1:A20"].FormatString = null;
            Assert.All(Enumerable.Range(0, 20), r => Assert.Equal("General", ws.GetCellAt(r, 0).FormatString));
        }

        [Fact]
        public void RangeFormatString_CreatesTheCellsItFormats()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["B2:C3"].FormatString = "0%";
            Assert.Equal("0%", ws["B2:C3"].FormatString);
            Assert.Equal(4, ws.FilledCells.Count);
        }

        // ---- first and last cell ----

        [Fact]
        public void FirstAndLastFilledCell_AreFoundInOrder_StyleOnlyCellsIncluded()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            Assert.Null(ws.FirstFilledCell);
            Assert.Null(ws.LastFilledCell);
            ws["D5"].Value = 1; ws["B9"].Value = 2; ws["F2"].Value = 3; ws["A12"].FormatString = "0.0"; ws["C12"].Value = "x";
            Assert.Equal("F2", ws.FirstFilledCell.AddressString);
            Assert.Equal("C12", ws.LastFilledCell.AddressString);
            ws["H12"].FormatString = "0"; // formatting alone makes a cell
            Assert.Equal("H12", ws.LastFilledCell.AddressString);
        }

        [Fact]
        public void FirstAndLastFilledCell_AgreeWithSorting_OnRandomSheets()
        {
            var rnd = new Random(99);
            for (var round = 0; round < 25; round++)
            {
                var ws = WorkBook.Create().DefaultWorkSheet;
                for (var i = rnd.Next(1, 200); i > 0; i--)
                {
                    var cell = ws.GetCellAt(rnd.Next(0, 100), rnd.Next(0, 30));
                    if (rnd.Next(4) == 0) cell.FormatString = "0.0"; else cell.Value = i;
                }
                var cells = ws.FilledCells;
                var first = cells.OrderBy(c => c.RowIndex).ThenBy(c => c.ColumnIndex).First();
                var last = cells.OrderByDescending(c => c.RowIndex).ThenByDescending(c => c.ColumnIndex).First();
                Assert.Equal(first.AddressString, ws.FirstFilledCell.AddressString);
                Assert.Equal(last.AddressString, ws.LastFilledCell.AddressString);
            }
        }

        // ---- culture ----

        [Fact]
        public void Evaluation_UsesTheInvariantCulture_AndGivesTheThreadItsOwnBack()
        {
            var saved = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo("de-DE");
                var ws = WorkBook.Create().DefaultWorkSheet;
                ws["A1"].Value = 1234.5;
                ws["A1"].FormatString = "#,##0.00";
                ws["B1"].Formula = "=TEXT(1234.5,\"#,##0.00\")";
                Assert.Equal("1,234.50", ws["A1"].StringValue);
                Assert.Equal("1,234.50", (string)ws["B1"].Value);
                Assert.Equal("de-DE", CultureInfo.CurrentCulture.Name);
                Assert.False(Internal.Formulas.EvaluationScope.Active);
            }
            finally { CultureInfo.CurrentCulture = saved; }
        }

        [Fact]
        public void Evaluation_InsideTheInvariantCulture_LeavesItAlone()
        {
            var saved = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
                var ws = WorkBook.Create().DefaultWorkSheet;
                ws["A1"].Formula = "=1+1";
                Assert.Equal(2d, ws["A1"].Value);
                Assert.Same(CultureInfo.InvariantCulture, CultureInfo.CurrentCulture);
            }
            finally { CultureInfo.CurrentCulture = saved; }
        }

        // ---- reading streams ----

        sealed class TrickleStream : MemoryStream
        {
            public TrickleStream(byte[] data) : base(data) { }
            public override int Read(byte[] buffer, int offset, int count) => base.Read(buffer, offset, Math.Min(count, 7));
        }

        sealed class ForwardOnly : MemoryStream
        {
            public ForwardOnly(byte[] data) : base(data) { }
            public override bool CanSeek => false;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        }

        byte[] Workbook()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = "hello";
            wb.DefaultWorkSheet["B2"].Value = 42;
            return wb.ToBinary();
        }

        [Fact]
        public void Load_FromStreams_ReadsEverything_WhateverTheStreamDoes()
        {
            var bytes = Workbook();
            foreach (var stream in new Stream[] { new MemoryStream(bytes), new TrickleStream(bytes), new ForwardOnly(bytes) })
                using (stream)
                {
                    var ws = WorkBook.Load(stream).DefaultWorkSheet;
                    Assert.Equal("hello", ws["A1"].StringValue);
                    Assert.Equal(42d, ws["B2"].DoubleValue);
                }
        }

        [Fact]
        public void Load_FromAStream_StartsAtItsCurrentPosition()
        {
            var bytes = Workbook();
            var ms = new MemoryStream();
            ms.Write(new byte[] { 9, 9, 9 }, 0, 3);
            ms.Write(bytes, 0, bytes.Length);
            ms.Position = 3;
            Assert.Equal("hello", WorkBook.Load(ms).DefaultWorkSheet["A1"].StringValue);
        }

        // ---- saving ----

        [Fact]
        public void SaveWithFormulas_CachesTheResults_ForReadersThatDoNotCalculate()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            for (var r = 0; r < 30; r++) { ws.GetCellAt(r, 0).Value = r; ws.GetCellAt(r, 1).Formula = "=A" + (r + 1) + "*2"; }
            wb.SaveAs(Tmp("f.xlsx"));
            var cells = WorkBook.StreamRows(Tmp("f.xlsx")).ToList();
            Assert.Equal(30, cells.Count);
            Assert.All(cells, row => Assert.Equal((double)row[0].Value * 2, row[1].Value));
        }

        [Fact]
        public void EvaluateAll_ComputesEveryFormula_AndSkipsUnsupportedOnes()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = 2;
            ws["A2"].Formula = "=A1*3";
            ws["A3"].Formula = "=NOSUCHFUNCTION(A1)";
            ws["A4"].Formula = "=A2+1";
            wb.EvaluateAll();
            Assert.Equal(6d, ws["A2"].Value);
            Assert.Equal(7d, ws["A4"].Value);
            Assert.Equal("#NAME?", ws["A3"].Value);
        }
    }
}
