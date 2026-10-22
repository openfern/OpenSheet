using System;
using System.IO;
using System.Linq;
using Xunit;
using static OpenSheet.Tests.Oracle.FormulaHarness;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Replays the answers Excel gave for <see cref="FormulaBank"/> (recorded by <see cref="FormulaOracleTests"/>), so the check
    /// runs everywhere, on CI and on machines without Excel. Same rule as the live test: unexpected differences fail.
    /// </summary>
    public class FormulaGoldenTests
    {
        [Fact]
        public void Evaluator_AgreesWithTheRecordedExcelAnswers()
        {
            var path = GoldenFile.OutputPath();
            Assert.True(File.Exists(path), "The recorded answers are missing: " + path);
            var lines = File.ReadAllLines(path).Where(l => l.Contains('\t')).Select(l => l.Split('\t')).ToList();
            Assert.True(lines.Count > 300, "The recorded answers look incomplete (" + lines.Count + " lines).");

            var formulas = lines.Select(l => l[0]).ToArray();
            var recorded = lines.Select(l => l[1] == "EXCEL-REJECTS" ? null : l[1]).ToArray();
            var ours = EvaluateWithOpenSheet(formulas, out _);

            var (unexpected, stale, _) = Judge(formulas, ours, recorded);
            var nl = Environment.NewLine + "  ";
            Assert.True(unexpected.Count == 0 && stale.Count == 0,
                (unexpected.Count > 0 ? "Differences from the recorded Excel answers:" + nl + string.Join(nl, unexpected) + Environment.NewLine : "") +
                (stale.Count > 0 ? "No longer different (remove from Known):" + nl + string.Join(nl, stale) : ""));
        }
    }
}
