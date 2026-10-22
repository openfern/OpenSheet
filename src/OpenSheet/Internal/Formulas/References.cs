using System;
using System.Text;
using System.Text.RegularExpressions;

namespace OpenSheet.Internal.Formulas
{
    /// <summary>
    /// Rewrites the cell references inside formula text the way Excel does when cells are copied or when rows and columns are
    /// inserted or deleted. Text inside "quotes", function names and defined names are never touched.
    /// </summary>
    internal static class References
    {
        public const int XlsxMaxRow = 1048575, XlsxMaxColumn = 16383;
        public const int XlsMaxRow = 65535, XlsMaxColumn = 255;

        // [sheet!] then A1 | A1:B2 | A:B | 1:2, not glued to a longer word and not a function call.
        static readonly Regex Token = new Regex(
            @"(?<![\w.$'\]!])" +
            @"(?<sheet>(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)!)?" +
            @"(?:(?<c1>\$?[A-Za-z]{1,3})(?<r1>\$?[0-9]+)(?::(?<c2>\$?[A-Za-z]{1,3})(?<r2>\$?[0-9]+))?" +
            @"|(?<wc1>\$?[A-Za-z]{1,3}):(?<wc2>\$?[A-Za-z]{1,3})" +
            @"|(?<wr1>\$?[0-9]+):(?<wr2>\$?[0-9]+))" +
            @"(?![\w(\[!.])",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        static readonly Regex StringLiteral = new Regex("\"(?:[^\"]|\"\")*\"", RegexOptions.CultureInvariant | RegexOptions.Compiled);

        // ---- public operations ----

        /// <summary>Copy semantics: every relative reference moves by the offset, absolute ($) parts stay. Off the sheet → #REF!.</summary>
        public static string Translate(string formula, int rowOffset, int columnOffset, int maxRow = XlsxMaxRow, int maxColumn = XlsxMaxColumn)
        {
            if (string.IsNullOrEmpty(formula) || (rowOffset == 0 && columnOffset == 0)) return formula;
            return Rewrite(formula, m => RewriteRef(m, c => Move(c, columnOffset, maxColumn), r => Move(r, rowOffset, maxRow)));
        }

        /// <summary>
        /// Structure semantics: rows (or columns) were inserted before index <paramref name="at"/> (<paramref name="delta"/> &gt; 0) or
        /// <c>-delta</c> of them were deleted starting at <paramref name="at"/>. Only references to <paramref name="targetSheet"/> change.
        /// <paramref name="ownSheet"/> is the sheet holding the formula (for references written without a sheet); null for names.
        /// </summary>
        public static string Adjust(string formula, string ownSheet, string targetSheet, bool rows, int at, int delta, int maxRow = XlsxMaxRow, int maxColumn = XlsxMaxColumn)
        {
            if (string.IsNullOrEmpty(formula) || delta == 0) return formula;
            var limit = rows ? maxRow : maxColumn;
            return Rewrite(formula, m =>
            {
                var sheet = SheetName(m.Groups["sheet"].Value) ?? ownSheet;
                if (sheet == null || !string.Equals(sheet, targetSheet, StringComparison.OrdinalIgnoreCase)) return null;
                return RewriteStructure(m, rows, at, delta, limit);
            });
        }

        // ---- engine ----

        static string Rewrite(string formula, Func<Match, string> onReference)
        {
            var sb = new StringBuilder(formula.Length + 8);
            var pos = 0;
            foreach (Match lit in StringLiteral.Matches(formula))
            {
                sb.Append(RewritePlain(formula.Substring(pos, lit.Index - pos), onReference));
                sb.Append(lit.Value);
                pos = lit.Index + lit.Length;
            }
            sb.Append(RewritePlain(formula.Substring(pos), onReference));
            return sb.ToString();
        }

        static string RewritePlain(string text, Func<Match, string> onReference)
        {
            if (text.Length == 0) return text;
            return Token.Replace(text, m => onReference(m) ?? m.Value);
        }

        static string SheetName(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return null;
            raw = raw.Substring(0, raw.Length - 1); // the "!"
            return raw.Length > 1 && raw[0] == '\'' ? raw.Substring(1, raw.Length - 2).Replace("''", "'") : raw;
        }

        static bool Column(string letters, out int index, out bool absolute)
        {
            absolute = letters.StartsWith("$", StringComparison.Ordinal);
            var s = absolute ? letters.Substring(1) : letters;
            index = 0;
            foreach (var ch in s.ToUpperInvariant()) index = index * 26 + (ch - 'A' + 1);
            index--;
            return index >= 0 && index <= XlsxMaxColumn;
        }

        static bool Row(string digits, out int index, out bool absolute)
        {
            absolute = digits.StartsWith("$", StringComparison.Ordinal);
            var s = absolute ? digits.Substring(1) : digits;
            if (!int.TryParse(s, out index) || index < 1 || index > XlsxMaxRow + 1) return false;
            index--;
            return true;
        }

        public static string ColumnName(int index)
        {
            var sb = new StringBuilder();
            for (index++; index > 0; index = (index - 1) / 26) sb.Insert(0, (char)('A' + (index - 1) % 26));
            return sb.ToString();
        }

        static string Part(bool absolute, string text) => (absolute ? "$" : "") + text;

        static int? Move(int x, int by, int max) => x + by < 0 || x + by > max ? (int?)null : x + by;

        // ---- translate ----

        static string RewriteRef(Match m, Func<int, int?> moveColumn, Func<int, int?> moveRow)
        {
            var sheet = m.Groups["sheet"].Value;
            string refError = sheet + "#REF!";
            if (m.Groups["c1"].Success)
            {
                if (!Column(m.Groups["c1"].Value, out var c1, out var ca1) || !Row(m.Groups["r1"].Value, out var r1, out var ra1)) return null;
                var a = Cell(c1, ca1, r1, ra1, moveColumn, moveRow);
                if (a == null) return refError;
                if (!m.Groups["c2"].Success) return sheet + a;
                if (!Column(m.Groups["c2"].Value, out var c2, out var ca2) || !Row(m.Groups["r2"].Value, out var r2, out var ra2)) return null;
                var b = Cell(c2, ca2, r2, ra2, moveColumn, moveRow);
                return b == null ? refError : sheet + a + ":" + b;
            }
            if (m.Groups["wc1"].Success)
            {
                if (!Column(m.Groups["wc1"].Value, out var c1, out var ca1) || !Column(m.Groups["wc2"].Value, out var c2, out var ca2)) return null;
                var x = ca1 ? (int?)c1 : moveColumn(c1); var y = ca2 ? (int?)c2 : moveColumn(c2);
                return x == null || y == null ? refError : sheet + Part(ca1, ColumnName(x.Value)) + ":" + Part(ca2, ColumnName(y.Value));
            }
            if (!Row(m.Groups["wr1"].Value, out var r1b, out var ra1b) || !Row(m.Groups["wr2"].Value, out var r2b, out var ra2b)) return null;
            var p = ra1b ? (int?)r1b : moveRow(r1b); var q = ra2b ? (int?)r2b : moveRow(r2b);
            return p == null || q == null ? refError : sheet + Part(ra1b, (p.Value + 1).ToString()) + ":" + Part(ra2b, (q.Value + 1).ToString());
        }

        static string Cell(int c, bool ca, int r, bool ra, Func<int, int?> moveColumn, Func<int, int?> moveRow)
        {
            var nc = ca ? c : moveColumn(c);
            var nr = ra ? r : moveRow(r);
            if (nc == null || nr == null) return null;
            return Part(ca, ColumnName(nc.Value)) + Part(ra, (nr.Value + 1).ToString());
        }

        // ---- structure ----

        static int? AdjustIndex(int x, int at, int delta, int max)
        {
            if (delta > 0) { var y = x >= at ? x + delta : x; return y > max ? (int?)null : y; }
            var n = -delta; var last = at + n - 1;
            if (x < at) return x;
            if (x <= last) return null;
            return x - n;
        }

        static bool AdjustSpan(int start, int end, int at, int delta, int max, out int newStart, out int newEnd)
        {
            if (delta > 0)
            {
                newStart = start >= at ? start + delta : start;
                newEnd = end >= at ? end + delta : end;
                if (newEnd > max) newEnd = max;
                return newStart <= max;
            }
            var n = -delta; var last = at + n - 1;
            newStart = start < at ? start : start <= last ? at : start - n;
            newEnd = end < at ? end : end <= last ? at - 1 : end - n;
            return newStart <= newEnd;
        }

        static string RewriteStructure(Match m, bool rows, int at, int delta, int limit)
        {
            var sheet = m.Groups["sheet"].Value;
            string refError = sheet + "#REF!";

            if (m.Groups["c1"].Success)
            {
                if (!Column(m.Groups["c1"].Value, out var c1, out var ca1) || !Row(m.Groups["r1"].Value, out var r1, out var ra1)) return null;
                if (!m.Groups["c2"].Success)
                {
                    var x = rows ? AdjustIndex(r1, at, delta, limit) : AdjustIndex(c1, at, delta, limit);
                    if (x == null) return refError;
                    return sheet + (rows ? Part(ca1, ColumnName(c1)) + Part(ra1, (x.Value + 1).ToString()) : Part(ca1, ColumnName(x.Value)) + Part(ra1, (r1 + 1).ToString()));
                }
                if (!Column(m.Groups["c2"].Value, out var c2, out var ca2) || !Row(m.Groups["r2"].Value, out var r2, out var ra2)) return null;
                int s, e;
                if (!AdjustSpan(rows ? r1 : c1, rows ? r2 : c2, at, delta, limit, out s, out e)) return refError;
                if (rows) return sheet + Part(ca1, ColumnName(c1)) + Part(ra1, (s + 1).ToString()) + ":" + Part(ca2, ColumnName(c2)) + Part(ra2, (e + 1).ToString());
                return sheet + Part(ca1, ColumnName(s)) + Part(ra1, (r1 + 1).ToString()) + ":" + Part(ca2, ColumnName(e)) + Part(ra2, (r2 + 1).ToString());
            }
            if (m.Groups["wc1"].Success)
            {
                if (rows) return null; // whole columns do not care about inserted rows
                if (!Column(m.Groups["wc1"].Value, out var c1, out var ca1) || !Column(m.Groups["wc2"].Value, out var c2, out var ca2)) return null;
                if (!AdjustSpan(c1, c2, at, delta, limit, out var s, out var e)) return refError;
                return sheet + Part(ca1, ColumnName(s)) + ":" + Part(ca2, ColumnName(e));
            }
            if (!rows) return null; // whole rows do not care about inserted columns
            if (!Row(m.Groups["wr1"].Value, out var wr1, out var wa1) || !Row(m.Groups["wr2"].Value, out var wr2, out var wa2)) return null;
            if (!AdjustSpan(wr1, wr2, at, delta, limit, out var rs, out var re)) return refError;
            return sheet + Part(wa1, (rs + 1).ToString()) + ":" + Part(wa2, (re + 1).ToString());
        }

        /// <summary>Applies <see cref="AdjustSpan"/> to a block of cells (a merged region, a validation range).</summary>
        public static bool AdjustBlock(ref int first, ref int last, int at, int delta, int max)
        {
            if (!AdjustSpan(first, last, at, delta, max, out var s, out var e)) return false;
            first = s; last = e;
            return true;
        }
    }
}
