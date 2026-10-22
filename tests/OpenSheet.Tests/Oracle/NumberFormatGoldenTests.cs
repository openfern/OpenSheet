using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Replays the text Excel drew for each value/format pair (recorded by <see cref="NumberFormatOracleTests"/>), so the display
    /// rules are checked everywhere, including on machines without Excel.
    /// </summary>
    public class NumberFormatGoldenTests
    {
        [Fact]
        public void FormattedCellValue_MatchesTheRecordedExcelText()
        {
            var path = NumberFormatGolden.OutputPath();
            Assert.True(File.Exists(path), "The recorded answers are missing: " + path);
            var rows = File.ReadAllLines(path).Select(l => l.Split('\t')).Where(p => p.Length == 3).ToList();
            Assert.True(rows.Count > 1000, "The recorded answers look incomplete (" + rows.Count + " lines).");

            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            var unexpected = new List<string>();
            var known = 0;
            for (var i = 0; i < rows.Count; i++)
            {
                var value = NumberFormatOracleTests.Decode(rows[i][0]);
                var format = rows[i][1]; var excel = rows[i][2];
                var cell = ws.GetCellAt(i, 0);
                cell.Value = value;
                cell.FormatString = format;
                string ours;
                try { ours = cell.FormattedCellValue; } catch (Exception e) { ours = "THROWS:" + e.GetType().Name; }
                if (ours == excel) continue;
                if (NumberFormatOracleTests.KnownReason(value, format, ours, excel) != null) { known++; continue; }
                unexpected.Add($"{NumberFormatOracleTests.Show(value),-18} [{format}]   OpenSheet=[{ours}]   Excel=[{excel}]");
            }
            Assert.True(unexpected.Count == 0, unexpected.Count + " unexpected difference(s) from the recorded Excel text (" + known + " known):" + Environment.NewLine + "  " +
                string.Join(Environment.NewLine + "  ", unexpected.Take(20)));
        }
    }
}
