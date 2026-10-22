using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OpenSheet.Tests
{
    public class StreamingTests : TestBase
    {
        string Build()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "id"; ws["B1"].Value = "name"; ws["D1"].Value = "when";
            ws["A2"].Value = 1; ws["B2"].Value = "ann"; ws["D2"].Value = new DateTime(2024, 3, 4);
            ws["A3"].Value = 2; ws["C3"].Value = true;
            ws["A4"].Formula = "=A2+A3";
            ws["B5"].Value = "tail";
            wb.CreateWorkSheet("Other")["A1"].Value = "x";
            var p = Tmp("s.xlsx");
            wb.SaveAs(p);
            return p;
        }

        [Fact]
        public void Rows_AreSparse_AndTyped()
        {
            var rows = WorkBook.StreamRows(Build()).ToList();
            Assert.Equal(new[] { 0, 1, 2, 3, 4 }, rows.Select(r => r.RowNumber).ToArray());
            Assert.Null(rows[0].Header);

            var r2 = rows[1];
            Assert.Equal(StreamCellType.Number, r2[0].Type);
            Assert.Equal(1d, r2[0].Value);
            Assert.Equal("ann", r2[1].Value);
            Assert.Null(r2[2]);                                  // gap
            Assert.Equal(StreamCellType.Date, r2[3].Type);
            Assert.Equal(new DateTime(2024, 3, 4), r2[3].Value);
            Assert.Equal("D", r2[3].ColumnLetter);

            Assert.Equal(StreamCellType.Boolean, rows[2][2].Type);
            Assert.True(rows[3][0].IsFormula);
            Assert.Equal(3d, rows[3][0].Value);                  // cached result of A2+A3 is only present after evaluation
        }

        [Fact]
        public void Header_NamesColumns_AndSkipsEarlierRows()
        {
            var rows = WorkBook.StreamRows(Build(), new StreamReadOptions { HasHeaderRow = true }).ToList();
            Assert.Equal(4, rows.Count);                          // header row consumed
            Assert.Equal(new[] { "id", "name", null, "when" }, rows[0].Header.ToArray());
            Assert.Equal("ann", rows[0]["name"].Value);
            Assert.Null(rows[1]["name"]);                         // known column, gap in this row
            Assert.True(rows[1].TryGetCell("id", out var id) && (double)id.Value == 2);
            Assert.False(rows[1].TryGetCell("nope", out _));
            Assert.Throws<KeyNotFoundException>(() => rows[0]["nope"]);
            Assert.Throws<InvalidOperationException>(() => WorkBook.StreamRows(Build()).First().TryGetCell("id", out _));
        }

        [Fact]
        public void Sheet_BySelection_AndNames()
        {
            var p = Build();
            Assert.Equal(new[] { "Sheet1", "Other" }, WorkBook.GetSheetNames(p).ToArray());
            Assert.Equal("x", WorkBook.StreamRows(p, "other").Single()[0].Value);
            Assert.Equal("x", WorkBook.StreamRows(p, 1).Single()[0].Value);
            Assert.Throws<ArgumentException>(() => WorkBook.StreamRows(p, "missing").ToList());
            Assert.Throws<ArgumentOutOfRangeException>(() => WorkBook.StreamRows(p, 5).ToList());
        }

        [Fact]
        public void FromStream_And_Bytes_LeaveCallerStreamOpen()
        {
            var p = Build();
            using (var fs = File.OpenRead(p))
            {
                Assert.Equal(5, WorkBook.StreamRows(fs).Count());
                Assert.True(fs.CanRead);
            }
            Assert.Equal(5, WorkBook.StreamRows(File.ReadAllBytes(p)).Count());
            Assert.Equal(2, WorkBook.GetSheetNames(File.ReadAllBytes(p)).Count);
        }

        [Fact]
        public void IsLazy_And_RejectsNonXlsx()
        {
            var seq = WorkBook.StreamRows(Tmp("not-created-yet.xlsx"));    // must not throw until enumerated
            Assert.Throws<FileNotFoundException>(() => seq.ToList());

            var xls = Tmp("o.xls");
            WorkBook.Create(ExcelFileFormat.XLS).SaveAs(xls);
            Assert.Throws<NotSupportedException>(() => WorkBook.StreamRows(xls).ToList());
        }

        [Fact]
        public void LargeSheet_Streams()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            for (var r = 0; r < 3000; r++) { ws.GetCellAt(r, 0).Value = r; ws.GetCellAt(r, 1).Value = "row" + r; }
            wb.SaveAs(Tmp("big.xlsx"));
            long sum = 0; var n = 0;
            foreach (var row in WorkBook.StreamRows(Tmp("big.xlsx"))) { sum += (long)(double)row[0].Value; n++; }
            Assert.Equal(3000, n);
            Assert.Equal(2999L * 3000 / 2, sum);
        }
    }
}
