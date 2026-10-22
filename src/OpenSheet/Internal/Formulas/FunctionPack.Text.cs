using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using NPOI.SS.Formula.Eval;

namespace OpenSheet.Internal.Formulas
{
    static partial class FunctionPack
    {
        static void DefineText()
        {
            Add("UNICHAR", (c, a) =>
            {
                Need(a, 1, 1);
                var n = (int)Math.Truncate(c.Num(a[0]));
                if (n <= 0 || n > 0x10FFFF || (n >= 0xD800 && n <= 0xDFFF)) throw XlError.Value;
                return S(char.ConvertFromUtf32(n));
            });
            Add("UNICODE", (c, a) =>
            {
                Need(a, 1, 1);
                var s = c.Str(a[0]);
                if (s.Length == 0) throw XlError.Value;
                return N(char.ConvertToUtf32(s, 0));
            });
            Add("LENB", (c, a) => { Need(a, 1, 1); return N(c.Str(a[0]).Length); });
            Add("ASC", (c, a) => { Need(a, 1, 1); return S(c.Str(a[0])); });
            Add("JIS", (c, a) => { Need(a, 1, 1); return S(c.Str(a[0])); });
            Add("TEXTBEFORE", (c, a) => TextSplitAt(c, a, before: true));
            Add("TEXTAFTER", (c, a) => TextSplitAt(c, a, before: false));
            Add("ARRAYTOTEXT", (c, a) =>
            {
                Need(a, 1, 2);
                if (a[0] is ErrorEval first) throw new XlError(first);   // an error in the array wins over a wrong format
                var format = c.Int(a, 1, 0);
                if (format != 0 && format != 1) throw XlError.Value;
                if (a[0] is RefEval && c.Scalar(a[0]) is BlankEval) throw XlError.Value;   // a lone empty cell is refused
                var strict = format == 1;
                var g = c.Area(a[0]);
                var sb = new StringBuilder();
                // concise: every value in reading order, ", " between them (a localized Excel uses its list separator instead); strict: an array constant, {1,2;3,4}
                for (var r = 0; r < g.Rows; r++)
                    for (var col = 0; col < g.Columns; col++)
                    {
                        if (r > 0 || col > 0) sb.Append(strict ? (col == 0 ? ";" : ",") : ", ");
                        sb.Append(CellText(g.Cells[r, col], strict));
                    }
                return S(strict ? "{" + sb + "}" : sb.ToString());
            });
            Add("VALUETOTEXT", (c, a) =>
            {
                Need(a, 1, 2);
                var format = c.Int(a, 1, 0);
                if (format != 0 && format != 1) throw XlError.Value;
                ValueEval v;
                try { v = c.Scalar(a[0]); } catch (XlError e) { v = e.Error; }   // an error is shown as its text, it does not spread
                return S(CellText(v, format == 1));
            });
        }

        static string CellText(ValueEval v, bool strict)
        {
            switch (v)
            {
                case StringEval s: return strict ? "\"" + s.StringValue.Replace("\"", "\"\"") + "\"" : s.StringValue;
                case NumberEval n: return Ctx.NumberText(n.NumberValue);
                case BoolEval b: return b.BooleanValue ? "TRUE" : "FALSE";
                case ErrorEval e: return e.ErrorString;
                default: return "";
            }
        }

        // TEXTBEFORE / TEXTAFTER(text, delimiter, [instance], [match_mode], [match_end], [if_not_found])
        static ValueEval TextSplitAt(Ctx c, ValueEval[] a, bool before)
        {
            Need(a, 2, 6);
            var text = c.Str(a[0]); var delim = c.Str(a[1]);
            var instance = c.Int(a, 2, 1); var ignoreCase = c.Int(a, 3, 0) == 1; var matchEnd = c.Int(a, 4, 0) == 1;
            if (instance == 0 || Math.Abs(instance) > text.Length + 1) throw XlError.Value;
            var cmp = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            ValueEval NotFound() { if (a.Length > 5 && !Ctx.IsMissing(a[5])) return a[5]; throw XlError.NA; }

            if (delim.Length == 0)
                return S(before == (instance > 0) ? "" : text);

            // positions of every delimiter, counted from the start or from the end
            var hits = new List<int>();
            for (var i = text.IndexOf(delim, 0, cmp); i >= 0; i = i + delim.Length <= text.Length ? text.IndexOf(delim, i + delim.Length, cmp) : -1) hits.Add(i);
            int pos, len = delim.Length;
            if (instance > 0)
            {
                if (instance <= hits.Count) pos = hits[instance - 1];
                else if (matchEnd && instance == hits.Count + 1) { pos = text.Length; len = 0; }
                else return NotFound();
            }
            else
            {
                var k = -instance;
                if (k <= hits.Count) pos = hits[hits.Count - k];
                else if (matchEnd && k == hits.Count + 1) { pos = 0; len = 0; }
                else return NotFound();
            }
            return S(before ? text.Substring(0, pos) : text.Substring(Math.Min(text.Length, pos + len)));
        }
    }
}
