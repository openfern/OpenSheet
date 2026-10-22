using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;
using Xunit.Abstractions;
using static OpenSheet.Tests.Oracle.FormulaHarness;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Runs every formula of <see cref="FormulaBank"/> through OpenSheet and through Excel and compares the results.
    /// The report lists what differs, so the evaluator's coverage is measured instead of assumed.
    /// Set OPENSHEET_RECORD_GOLDEN=1 to write Excel's answers into <c>formula-bank.golden.txt</c>, which
    /// <see cref="FormulaGoldenTests"/> replays on machines without Excel.
    /// </summary>
    [Collection("Excel")]
    public class FormulaOracleTests
    {
        readonly ExcelOracle _excel;
        readonly ITestOutputHelper _out;
        public FormulaOracleTests(ExcelOracle excel, ITestOutputHelper output) { _excel = excel; _out = output; }

        [ExcelFact]
        public void Evaluator_AgreesWithExcel_OnTheFormulaBank()
        {
            var formulas = FormulaBank.Formulas.Distinct().ToArray();
            var ours = EvaluateWithOpenSheet(formulas, out var ws);

            var theirs = new string[formulas.Length];
            using (var doc = _excel.New())
            {
                dynamic xs = doc.Sheet(1); xs.Name = "S";
                FillExcel(xs);
                for (var i = 0; i < formulas.Length; i++)
                {
                    try { xs.Cells(i * Stride + 1, FormulaColumn + 1).Formula2 = formulas[i]; }
                    catch (Exception) { theirs[i] = "EXCEL-REJECTS"; }
                }
                _excel.App.Calculate();
                for (var i = 0; i < formulas.Length; i++)
                {
                    if (theirs[i] != null) continue;
                    try { theirs[i] = Canon(xs.Cells(i * Stride + 1, FormulaColumn + 1).Value2); }
                    catch (Exception e) { theirs[i] = "EXCEL-READ-FAILS:" + e.Message.Split('\n')[0]; }
                }
            }

            var report = new StringBuilder();
            var usable = theirs.Select(t => t == "EXCEL-REJECTS" || t.StartsWith("EXCEL-READ") ? null : t).ToArray();
            for (var i = 0; i < formulas.Length; i++) if (usable[i] == null) report.AppendLine($"SKIP   {formulas[i]}   (Excel: {theirs[i]})");
            var (unexpected, stale, all) = Judge(formulas, ours, usable);
            var comparable = usable.Count(t => t != null);
            report.AppendLine($"compared {comparable} formulas, {comparable - all.Count} agree, {all.Count} differ");
            foreach (var m in all) report.AppendLine("DIFF   " + m);

            var path = Path.Combine(Path.GetTempPath(), "opensheet-formula-oracle.txt");
            File.WriteAllText(path, report.ToString());
            _out.WriteLine(report.ToString());

            if (Environment.GetEnvironmentVariable("OPENSHEET_RECORD_GOLDEN") == "1")
            {
                // every formula is recorded, in order, because some answers depend on where the formula sits (ROW, COLUMN)
                var golden = Enumerable.Range(0, formulas.Length).Select(i => formulas[i] + "\t" + (usable[i] ?? "EXCEL-REJECTS")).ToList();
                File.WriteAllLines(GoldenFile.SourcePath(), golden, new UTF8Encoding(false));
                _out.WriteLine("recorded " + golden.Count + " answers into " + GoldenFile.SourcePath());
            }

            var nl = Environment.NewLine + "  ";
            Assert.True(unexpected.Count == 0 && stale.Count == 0,
                (unexpected.Count > 0 ? "New differences from Excel:" + nl + string.Join(nl, unexpected) + Environment.NewLine : "") +
                (stale.Count > 0 ? "No longer different (remove from Known):" + nl + string.Join(nl, stale) : ""));
        }
    }

    static class GoldenFile
    {
        public const string Name = "formula-bank.golden.txt";

        /// <summary>The copy next to the test sources (what gets committed), found by walking up from the test binaries.</summary>
        public static string SourcePath()
        {
            for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
                if (File.Exists(Path.Combine(dir.FullName, "OpenSheet.Tests.csproj"))) return Path.Combine(dir.FullName, "Oracle", Name);
            throw new InvalidOperationException("Cannot find the test project folder.");
        }

        /// <summary>The copy that was built into the test output.</summary>
        public static string OutputPath() => Path.Combine(AppContext.BaseDirectory, "Oracle", Name);
    }
}
