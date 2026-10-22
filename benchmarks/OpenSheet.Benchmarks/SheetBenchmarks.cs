using System;
using System.IO;
using System.Linq;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace OpenSheet.Benchmarks
{
    /// <summary>Public-API workloads only, so the same file measures any build of the library.</summary>
    [MemoryDiagnoser]
    public class SheetBenchmarks
    {
        const int Rows = 20_000;
        const int Columns = 8;

        string _dir;
        string _xlsxPath;
        string _csvPath;
        WorkBook _book;
        WorkBook _formulas;
        WorkBook _scratch;

        [GlobalSetup]
        public void Setup()
        {
            _dir = Path.Combine(Path.GetTempPath(), "opensheet-bench-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);

            _book = Fill(WorkBook.Create(), Rows);
            _xlsxPath = Path.Combine(_dir, "data.xlsx");
            _book.SaveAs(_xlsxPath);

            _csvPath = Path.Combine(_dir, "data.csv");
            _book.SaveAs(_csvPath);

            _formulas = Fill(WorkBook.Create(), 5_000);
            var ws = _formulas.DefaultWorkSheet;
            for (var r = 1; r <= 5_000; r++) ws["I" + (r + 1)].Formula = "=SUM(A" + (r + 1) + ",C" + (r + 1) + ")*2";

            _scratch = Fill(WorkBook.Create(), 5_000);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            try { Directory.Delete(_dir, true); } catch (IOException) { }
        }

        static WorkBook Fill(WorkBook book, int rows)
        {
            var ws = book.DefaultWorkSheet;
            string[] head = { "Id", "Name", "Amount", "When", "Active", "Note", "Ratio", "Code" };
            for (var c = 0; c < head.Length; c++) ws.GetCellAt(0, c).Value = head[c];
            var day = new DateTime(2024, 1, 1);
            for (var r = 1; r <= rows; r++)
            {
                ws.GetCellAt(r, 0).Value = r;
                ws.GetCellAt(r, 1).Value = "Customer " + r;
                ws.GetCellAt(r, 2).Value = r * 1.25;
                ws.GetCellAt(r, 3).Value = day.AddDays(r % 365);
                ws.GetCellAt(r, 4).Value = r % 2 == 0;
                ws.GetCellAt(r, 5).Value = "say \"hi\", then <leave> & go " + r;
                ws.GetCellAt(r, 6).Value = (r % 100) / 100.0;
                ws.GetCellAt(r, 7).Value = "C" + (r % 977).ToString("D4");
            }
            return book;
        }

        // ---- text exports ----

        [Benchmark] public MemoryStream Export_Csv() => _book.ToCsvStream();
        [Benchmark] public MemoryStream Export_Json() => _book.ToJsonStream();
        [Benchmark] public MemoryStream Export_Xml() => _book.ToXmlStream();
        [Benchmark] public string Export_Html() => _book.ExportToHtmlString();

        // ---- streaming read ----

        [Benchmark]
        public double Stream_Rows()
        {
            double sum = 0;
            foreach (var row in WorkBook.StreamRows(_xlsxPath)) if (row[0]?.Value is double d) sum += d;
            return sum;
        }

        [Benchmark]
        public double Stream_RowsWithHeader()
        {
            double sum = 0;
            foreach (var row in WorkBook.StreamRows(_xlsxPath, new StreamReadOptions { HasHeaderRow = true }))
                if (row["Amount"]?.Value is double d) sum += d;
            return sum;
        }

        // ---- text import ----

        [Benchmark] public WorkBook Import_Csv() => WorkBook.LoadCSV(_csvPath);

        // ---- cell access ----

        [Benchmark]
        public WorkBook Write_ByAddress()
        {
            var book = WorkBook.Create();
            var ws = book.DefaultWorkSheet;
            for (var r = 1; r <= 5_000; r++)
            {
                ws["A" + r].Value = r;
                ws["B" + r].Value = "x";
                ws["C" + r].Value = r * 0.5;
            }
            return book;
        }

        [Benchmark] public string Read_RangeStringValue() => _book.DefaultWorkSheet["A1:H2000"].StringValue;
        [Benchmark] public decimal Read_RangeSum() => _book.DefaultWorkSheet["A1:H20000"].Sum();
        [Benchmark] public int Read_AddressStrings() => _book.DefaultWorkSheet["A1:H5000"].Sum(c => c.AddressString.Length);
        [Benchmark] public System.Data.DataTable Read_ToDataTable() => _book.DefaultWorkSheet.ToDataTable(true);

        [Benchmark]
        public void Edit_FormatString()
        {
            _scratch.DefaultWorkSheet["A2:H3000"].FormatString = "#,##0.00";
            _scratch.DefaultWorkSheet["A2:H3000"].FormatString = "0.0";
        }

        // ---- saving ----

        [Benchmark]
        public byte[] Save_XlsxWithFormulas()
        {
            _formulas.DefaultWorkSheet["A1"].Value = "Id"; // marks the workbook edited, so formula results are cached again
            return _formulas.ToBinary();
        }
    }
}
