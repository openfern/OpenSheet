using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>Shared by the live Excel comparison and by the test that replays Excel's recorded answers.</summary>
    static class FormulaHarness
    {
        public const int Stride = 12;          // rows between formulas, so dynamic-array results can spill without hitting the next one
        public const int FormulaColumn = 12;   // M

        public static readonly string[] Words = { "apple", "Banana", "cherry", "date", "Elder", "fig", "grape", "honey" };
        public static readonly double[] Nums = { 3, 1, 4, 1, 5, 9, 2, 6 };

        public static void Fill(WorkSheet ws)
        {
            for (var i = 0; i < 8; i++) { ws.GetCellAt(i, 0).Value = Nums[i]; ws.GetCellAt(i, 1).Value = Words[i]; }
            for (var i = 0; i < 4; i++)
            {
                ws.GetCellAt(i, 2).Value = 10.0 * (i + 1);
                ws.GetCellAt(i, 3).Value = new[] { "x", "y", "x", "z" }[i];
                ws.GetCellAt(i, 7).Value = (double)(i + 1);
                ws.GetCellAt(i, 8).Value = new[] { "one", "two", "three", "four" }[i];
                ws.GetCellAt(i, 9).Value = 10.5 + 10 * i;
            }
            ws.GetCellAt(0, 4).Value = 45000.0;
            ws.GetCellAt(0, 5).Value = " padded   text ";
            ws.GetCellAt(1, 5).Value = "Hello, World";
            ws.GetCellAt(2, 5).Value = "2024-03-09";
            ws.GetCellAt(3, 5).Value = "1,234.5";
            ws.GetCellAt(0, 10).Formula = "=NA()"; ws.GetCellAt(1, 10).Value = 1.0; ws.GetCellAt(2, 10).Value = 2.0; ws.GetCellAt(3, 10).Formula = "=1/0";   // K1:K4: two errors among numbers
            ws.GetCellAt(0, 11).Formula = "=A1*2+A2^2";   // a formula cell for ISFORMULA / FORMULATEXT (no function names: Excel writes them in its own language)
        }

        public static void FillExcel(dynamic ws)
        {
            for (var i = 0; i < 8; i++) { ws.Cells(i + 1, 1).Value2 = Nums[i]; ws.Cells(i + 1, 2).Value2 = Words[i]; }
            for (var i = 0; i < 4; i++)
            {
                ws.Cells(i + 1, 3).Value2 = 10.0 * (i + 1);
                ws.Cells(i + 1, 4).Value2 = new[] { "x", "y", "x", "z" }[i];
                ws.Cells(i + 1, 8).Value2 = (double)(i + 1);
                ws.Cells(i + 1, 9).Value2 = new[] { "one", "two", "three", "four" }[i];
                ws.Cells(i + 1, 10).Value2 = 10.5 + 10 * i;
            }
            ws.Cells(1, 5).Value2 = 45000.0;
            ws.Cells(1, 6).Value2 = " padded   text ";
            ws.Cells(2, 6).Value2 = "Hello, World";
            ws.Cells(3, 6).NumberFormat = "@"; ws.Cells(3, 6).Value2 = "2024-03-09";
            ws.Cells(4, 6).NumberFormat = "@"; ws.Cells(4, 6).Value2 = "1,234.5";
            ws.Cells(1, 11).Formula = "=NA()"; ws.Cells(2, 11).Value2 = 1.0; ws.Cells(3, 11).Value2 = 2.0; ws.Cells(4, 11).Formula = "=1/0";
            ws.Cells(1, 12).Formula = "=A1*2+A2^2";
        }

        /// <summary>Evaluates each formula in a fresh OpenSheet workbook holding the shared data block.</summary>
        public static string[] EvaluateWithOpenSheet(string[] formulas, out WorkSheet sheet)
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet; ws.Name = "S";
            Fill(ws);
            var ours = new string[formulas.Length];
            for (var i = 0; i < formulas.Length; i++)
            {
                try
                {
                    var cell = ws.GetCellAt(i * Stride, FormulaColumn);
                    cell.Formula = formulas[i];
                    ours[i] = Canon(cell.Value);
                }
                catch (Exception e) { ours[i] = "THROWS:" + e.GetType().Name + ":" + e.Message.Split('\n')[0]; }
            }
            sheet = ws;
            return ours;
        }

        /// <summary>A result reduced to something comparable: numbers as text with 12 digits, everything tagged with its type.</summary>
        public static string Canon(object v)
        {
            switch (v)
            {
                case null: return "empty";
                case DateTime dt: return "num:" + Num(Internal.Serial.FromDateTime(dt));
                case double d: return "num:" + Num(d);
                case int i: return i < -2000000000 ? "err:" + Diffs.ExcelErrorText(i) : "num:" + Num(i);
                case bool b: return "bool:" + b;
                case string s: return s.StartsWith("#", StringComparison.Ordinal) && s.Length > 2 && (s.EndsWith("!") || s.EndsWith("?") || s.EndsWith("A")) ? "err:" + s : "str:" + s;
                default: return "other:" + v;
            }
        }

        static string Num(double d) => Math.Abs(d) < 1e-12 ? "0" : d.ToString("G12", CultureInfo.InvariantCulture);

        public static bool Equivalent(string a, string b, string formula)
        {
            if (a == b) return true;
            if (a.StartsWith("num:") && b.StartsWith("num:") &&
                double.TryParse(a.Substring(4), NumberStyles.Float, CultureInfo.InvariantCulture, out var x) &&
                double.TryParse(b.Substring(4), NumberStyles.Float, CultureInfo.InvariantCulture, out var y))
            {
                // Excel itself stops its iterative solvers at about 1e-7, so exact digits cannot be matched.
                var tolerance = formula.StartsWith("=RATE(") || formula.StartsWith("=XIRR(") || formula.StartsWith("=IRR(") ? 1e-5 : 1e-9;
                return Math.Abs(x - y) <= tolerance * Math.Max(1, Math.Max(Math.Abs(x), Math.Abs(y)));
            }
            return false;
        }

        /// <summary>
        /// Where OpenSheet knowingly differs from Excel, with the reason. A new difference fails the test; a difference that
        /// disappears makes the entry stale, which fails it too, so this list can only be edited deliberately.
        /// </summary>
        public static readonly Dictionary<string, string> Known = new Dictionary<string, string>
        {
            ["=-2^2"] = "NPOI's parser binds ^ tighter than unary minus; Excel binds unary minus first",
            ["=A1:A3"] = "a bare range outside an array context spills in Excel 365; NPOI has no dynamic arrays",
            ["=SUMPRODUCT((A1:A8>2)*(A1:A8<9))"] = "NPOI cannot do array arithmetic inside functions",
            ["=AND(A1:A3>0)"] = "NPOI cannot do array arithmetic inside functions",
            ["=XOR(A1:A3>1)"] = "NPOI cannot do array arithmetic inside functions",
            ["=TRANSPOSE(A1:A3)"] = "array result; NPOI has no dynamic arrays",
            ["=RANDARRAY(1)*0"] = "dynamic array function",
            ["=TEXTSPLIT(\"a,b\",\",\")"] = "dynamic array function",
            ["=SORT(A1:A4)"] = "dynamic array function",
            ["=UNIQUE(D1:D4)"] = "dynamic array function",
            ["=FILTER(A1:A8,A1:A8>4)"] = "dynamic array function",
            ["=SEQUENCE(3)"] = "dynamic array function",
            ["=ISOMITTED(1)"] = "only meaningful inside LAMBDA",
            ["=TEXT(0.5,\"# ?/?\")"] = "NPOI drops the leading blank of an empty whole-number part",
        };

        /// <summary>Compares two answer lists; returns the formulas that differ and are not in <see cref="Known"/>, and the stale entries.</summary>
        public static (List<string> unexpected, List<string> stale, List<string> all) Judge(string[] formulas, string[] ours, string[] excel)
        {
            var all = new List<string>();
            var differing = new HashSet<string>();
            for (var i = 0; i < formulas.Length; i++)
            {
                if (excel[i] == null) continue;
                if (Equivalent(ours[i], excel[i], formulas[i])) continue;
                differing.Add(formulas[i]);
                all.Add($"{formulas[i]}   OpenSheet=[{ours[i]}]   Excel=[{excel[i]}]");
            }
            var unexpected = all.Where(m => !Known.ContainsKey(m.Substring(0, m.IndexOf("   OpenSheet=", StringComparison.Ordinal)))).ToList();
            var stale = Known.Keys.Where(f => !differing.Contains(f) && formulas.Contains(f)).ToList();
            return (unexpected, stale, all);
        }
    }
}
