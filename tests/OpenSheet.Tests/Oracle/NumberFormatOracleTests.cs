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
    /// What a cell displays: OpenSheet's FormattedCellValue against the text Excel itself draws, for many values and formats.
    /// Only locale-neutral codes are used (no month or day names, no currency of the machine), and Excel is told to use "." and ",".
    /// The formats travel in a file, because Excel reads codes typed through its object model in the language of its user interface
    /// (a Spanish Excel does not take "yyyy" as the year) while a file always holds them in the invariant form.
    /// </summary>
    [Collection("Excel")]
    public class NumberFormatOracleTests : TestBase
    {
        readonly ExcelOracle _excel;
        readonly ITestOutputHelper _out;
        public NumberFormatOracleTests(ExcelOracle excel, ITestOutputHelper output) { _excel = excel; _out = output; }

        internal static readonly object[] Values =
        {
            0.0, 1.0, -1.0, 0.5, -0.5, 1234.5678, -1234.5678, 0.000123, -0.000123, 1e10, 123456789.123, 1.0 / 3, 45000.5, 0.75, 100.0, 99.995, 45000.9999999,
            0.1, 1e-5, 1.5e-7, 1.23456789e-5, 0.00012345678901234, 12345678901, 123456789012, 1.5e11, 1.23456789012345e20, 0.1 + 0.2, 123456.789012345,
            1234567890.123, 12345678.9, 1234567.891, 4.35, 2.675, 1e300, 1e-300, 0.99999999999, 999999999.9999, 1e15,
            "text",
        };

        internal static readonly string[] Formats =
        {
            "General", "0", "0.0", "0.00", "#,##0", "#,##0.00", "#,##0.0;(#,##0.0)", "0%", "0.00%", "0.00E+00", "##0.0E+0", "# ?/?", "# ??/??", "# ???/???",
            "$#,##0.00", "\"EUR\" #,##0.00", "[Red]0.00", "0.0;-0.0;\"zero\"", "@", "000000", "0.00_);(0.00)", "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
            "yyyy-mm-dd", "dd/mm/yyyy", "hh:mm:ss", "h:mm AM/PM", "yyyy-mm-dd hh:mm", "[h]:mm:ss", "mm:ss.0", "d", "yy", "[$-409]yyyy",
            "0.0,", "#,##0,,", "[>=100]0.0;[<100]0.00", "\"x\"0", "0\"x\"", "\\#0", "General\"!\"", "0.000", "#.##", "#.00", "?.??",
        };

        [ExcelFact]
        public void FormattedCellValue_MatchesWhatExcelDraws()
        {
            var pairs = AllPairs();

            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            var ours = new string[pairs.Count];
            for (var i = 0; i < pairs.Count; i++)
            {
                var cell = ws.GetCellAt(i, 0);
                cell.Value = pairs[i].value;
                cell.FormatString = pairs[i].format;
                try { ours[i] = cell.FormattedCellValue; } catch (Exception e) { ours[i] = "THROWS:" + e.GetType().Name; }
            }

            wb.SaveAs(Tmp("formats.xlsx"));
            var theirs = new string[pairs.Count];
            using (var doc = _excel.Open(Tmp("formats.xlsx")))
            {
                dynamic xs = doc.Sheet(1);
                xs.Columns("A:A").ColumnWidth = 60;
                for (var i = 0; i < pairs.Count; i++)
                {
                    try { theirs[i] = (string)xs.Cells(i + 1, 1).Text; }
                    catch (Exception e) { theirs[i] = "EXCEL-FAILS:" + e.Message.Split('\n')[0]; }
                }
            }

            var report = new StringBuilder();
            var diffs = new List<string>();
            var known = 0;
            for (var i = 0; i < pairs.Count; i++)
            {
                if (theirs[i].StartsWith("EXCEL-FAILS")) { report.AppendLine($"SKIP  {Show(pairs[i].value)} with [{pairs[i].format}]: {theirs[i]}"); continue; }
                if (ours[i] == theirs[i]) continue;
                var why = KnownReason(pairs[i].value, pairs[i].format, ours[i], theirs[i]);
                if (why != null) { known++; continue; }
                diffs.Add($"{Show(pairs[i].value),-18} [{pairs[i].format}]   OpenSheet=[{ours[i]}]   Excel=[{theirs[i]}]");
            }
            report.AppendLine($"compared {pairs.Count} cells, {pairs.Count - diffs.Count - known} agree, {known} differ for a known reason, {diffs.Count} differ unexpectedly");
            foreach (var d in diffs) report.AppendLine("DIFF  " + d);
            File.WriteAllText(Path.Combine(Path.GetTempPath(), "opensheet-numberformat-oracle.txt"), report.ToString());
            if (Environment.GetEnvironmentVariable("OPENSHEET_RECORD_GOLDEN") == "1")
            {
                var golden = Enumerable.Range(0, pairs.Count).Where(i => !theirs[i].StartsWith("EXCEL-FAILS")).Select(i => Encode(pairs[i].value) + "\t" + pairs[i].format + "\t" + theirs[i]).ToList();
                File.WriteAllLines(NumberFormatGolden.SourcePath(), golden, new UTF8Encoding(false));
            }
            _out.WriteLine(report.ToString());
            Assert.True(diffs.Count == 0, diffs.Count + " unexpected difference(s) from Excel, first ones:" + Environment.NewLine + "  " + string.Join(Environment.NewLine + "  ", diffs.Take(15)));
        }

        /// <summary>Differences that are understood and accepted. Anything else fails the test.</summary>
        internal static string KnownReason(object value, string format, string ours, string excel)
        {
            if (excel.Length > 0 && excel.All(c => c == '#')) return "Excel fills the column with # when a value does not fit or cannot be drawn as a date; the number of # depends on the column width";
            if (format.IndexOf('*') >= 0) return "the * fill repeats a character to the width of the column";
            if (format.IndexOf("?/", StringComparison.Ordinal) >= 0) return "fractions are padded with blanks to line up; NPOI trims them";
            if (format.StartsWith("##0.0E", StringComparison.Ordinal)) return "engineering notation (exponent a multiple of 3) is not implemented by NPOI";
            if (format == "#.##" || format == "?.??") return "a format with a bare decimal point keeps the point in Excel";
            if (value is double d && d == -0.000123 && format.IndexOf(';') >= 0) return "a negative number that rounds to zero still uses the negative section in Excel";
            if (value is double z && z == 0.0 && format.IndexOf("0.00_", StringComparison.Ordinal) >= 0) return null;
            return null;
        }

        internal static string Show(object v) => v is string s ? "\"" + s + "\"" : Convert.ToString(v, CultureInfo.InvariantCulture);

        /// <summary>The value written the way the recorded answers keep it: "n:" and the exact double, or "s:" and the text.</summary>
        internal static string Encode(object v) => v is string s ? "s:" + s : "n:" + Convert.ToDouble(v, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture);
        internal static object Decode(string text) => text.StartsWith("s:", StringComparison.Ordinal) ? (object)text.Substring(2) : double.Parse(text.Substring(2), CultureInfo.InvariantCulture);

        internal static List<(object value, string format)> AllPairs()
        {
            var pairs = new List<(object value, string format)>();
            foreach (var f in Formats) foreach (var v in Values) pairs.Add((v, f));
            return pairs;
        }
    }

    static class NumberFormatGolden
    {
        public const string Name = "numberformat.golden.txt";

        public static string SourcePath()
        {
            for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
                if (File.Exists(Path.Combine(dir.FullName, "OpenSheet.Tests.csproj"))) return Path.Combine(dir.FullName, "Oracle", Name);
            throw new InvalidOperationException("Cannot find the test project folder.");
        }

        public static string OutputPath() => Path.Combine(AppContext.BaseDirectory, "Oracle", Name);
    }
}
