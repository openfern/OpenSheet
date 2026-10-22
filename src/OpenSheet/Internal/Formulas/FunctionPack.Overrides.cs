using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;

namespace OpenSheet.Internal.Formulas
{
    /// <summary>
    /// True while OpenSheet is evaluating a formula on the current thread. Corrections to NPOI's built-in functions only apply
    /// inside it, so other code in the same process that uses NPOI directly keeps NPOI's own behavior.
    /// </summary>
    static class EvaluationScope
    {
        [ThreadStatic] static int _depth;
        public static bool Active => _depth > 0;
        public static void Enter() => _depth++;
        public static void Leave() => _depth--;
    }

    sealed class ScopedBuiltin : NPOI.SS.Formula.Functions.Function
    {
        readonly NPOI.SS.Formula.Functions.Function _original;
        readonly PackFunction _ours;
        public ScopedBuiltin(NPOI.SS.Formula.Functions.Function original, PackFunction ours) { _original = original; _ours = ours; }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) =>
            EvaluationScope.Active ? _ours.Evaluate(args, srcRowIndex, srcColumnIndex) : _original.Evaluate(args, srcRowIndex, srcColumnIndex);
    }

    sealed class ScopedAtp : FreeRefFunction
    {
        readonly FreeRefFunction _original;
        readonly PackFunction _ours;
        public ScopedAtp(FreeRefFunction original, PackFunction ours) { _original = original; _ours = ours; }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) =>
            EvaluationScope.Active || _original == null ? _ours.Evaluate(args, ec) : _original.Evaluate(args, ec);
    }

    static partial class FunctionPack
    {
        static readonly Dictionary<string, PackFunction> Corrections = new Dictionary<string, PackFunction>(StringComparer.OrdinalIgnoreCase);

        static void Correct(string name, Impl impl) => Corrections[name] = new PackFunction(impl);

        static void DefineCorrections()
        {
            DefineEdgeCorrections();
            Correct("FLOOR", (c, a) =>
            {
                Need(a, 2, 2);
                double n = c.Num(a[0]), s = c.Num(a[1]);
                if (n == 0) return N(0);
                if (s == 0) throw XlError.Div0;
                if (n > 0 && s < 0) throw XlError.Num;
                var r = n > 0 ? Math.Floor(n / s) * s : s > 0 ? -Math.Ceiling(-n / s) * s : -Math.Floor(-n / -s) * -s;
                return N(r);
            });
            Correct("TRIM", (c, a) =>
            {
                Need(a, 1, 1);
                return S(Regex.Replace(c.Str(a[0]), " {2,}", " ").Trim(' '));
            });
            Correct("CODE", (c, a) =>
            {
                Need(a, 1, 1);
                var s = c.Str(a[0]);
                if (s.Length == 0) throw XlError.Value;
                return N(s[0]);
            });
            Correct("SEARCH", (c, a) =>
            {
                Need(a, 2, 3);
                var find = c.Str(a[0]); var within = c.Str(a[1]); var start = c.Int(a, 2, 1);
                if (start < 1 || start > within.Length + 1) throw XlError.Value;
                var m = new Regex(WildcardToRegex(find), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Match(within, start - 1);
                if (!m.Success) throw XlError.Value;
                return N(m.Index + 1);
            });
            Correct("ROMAN", (c, a) =>
            {
                Need(a, 1, 2);
                var n = Math.Floor(c.Num(a[0]));
                var form = a.Length > 1 && !Ctx.IsMissing(a[1]) ? (a[1] is BoolEval b ? (b.BooleanValue ? 0 : 4) : (int)Math.Truncate(c.Num(a[1]))) : 0;
                if (n < 0 || n > 3999 || form < 0 || form > 4) throw XlError.Value;
                return S(Roman((int)n, form));
            });
            Correct("WORKDAY.INTL", WorkdayIntl);
            Corrections["NORM.S.DIST"] = Map["NORM.S.DIST"]; // NPOI's takes one argument; Excel's takes the cumulative flag too
            Corrections["CEILING.MATH"] = Map["CEILING.MATH"]; // NPOI's refuses a mode given as text ("1") where Excel reads the number
            Corrections["FLOOR.MATH"] = Map["FLOOR.MATH"];
            Corrections["POISSON"] = Map["POISSON"]; // NPOI reads a "cumulative" argument given as text
            // names NPOI implements itself: where it differs from Excel at the edges (booleans, blanks, text, limits) the checked version rules
            foreach (var name in new[] { "BETAINV", "BINOMDIST", "CHIDIST", "CHIINV", "CONFIDENCE", "EXPONDIST", "FDIST", "FINV", "GAMMADIST", "GAMMAINV",
                                         "LOGINV", "NORMDIST", "NORMINV", "NORMSINV", "WEIBULL", "NORM.DIST", "NORM.INV", "NORM.S.INV", "T.INV" })
                Corrections[name] = Map[name];
            Corrections["MODE"] = Map["MODE.SNGL"]; // NPOI refuses text typed into the formula, which Excel ignores
            Correct("XLOOKUP", XLookup);
            Correct("DEC2BIN", (c, a) => { RejectBooleans(a); return DecTo(c, a, 2, 10); });
            Correct("DEC2HEX", (c, a) => { RejectBooleans(a); return DecTo(c, a, 16, 10); });
            Correct("DEC2OCT", (c, a) => { RejectBooleans(a); return DecTo(c, a, 8, 10); });
        }

        // XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
        static ValueEval XLookup(Ctx c, ValueEval[] a)
        {
            Need(a, 3, 6);
            var look = c.Scalar(a[0]); var keys = c.Area(a[1]); var rows = c.Area(a[2]);
            var matchMode = c.Int(a, 4, 0); var searchMode = c.Int(a, 5, 1);
            if (matchMode < -1 || matchMode > 2 || Math.Abs(searchMode) < 1 || Math.Abs(searchMode) > 2) throw XlError.Value;
            var vertical = keys.Columns == 1;
            if (!vertical && keys.Rows != 1) throw XlError.Value;
            var n = vertical ? keys.Rows : keys.Columns;
            if ((vertical ? rows.Rows : rows.Columns) != n) throw XlError.Value;
            ValueEval Key(int i) => vertical ? keys.Cells[i, 0] : keys.Cells[0, i];

            var order = Enumerable.Range(0, n);
            if (searchMode < 0) order = order.Reverse();
            var best = -1;
            Regex wildcard = matchMode == 2 && look is StringEval ls ? new Regex("^" + WildcardToRegex(ls.StringValue) + "$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) : null;
            foreach (var i in order)
            {
                var k = Key(i);
                if (wildcard != null) { if (k is StringEval ks && wildcard.IsMatch(ks.StringValue)) { best = i; break; } continue; }
                var cmp = CompareValues(k, look);
                if (cmp == null) continue;
                if (cmp == 0) { best = i; break; }
                if (matchMode == -1 && cmp < 0 && (best < 0 || CompareValues(k, Key(best)) > 0)) best = i;
                if (matchMode == 1 && cmp > 0 && (best < 0 || CompareValues(k, Key(best)) < 0)) best = i;
            }
            if (best < 0)
            {
                if (a.Length > 3 && !Ctx.IsMissing(a[3])) return c.Scalar(a[3]);
                throw XlError.NA;
            }
            var hit = vertical ? rows.Cells[best, 0] : rows.Cells[0, best];
            return hit is BlankEval ? N(0) : hit;
        }

        // Excel's ordering within one type; values of different types never match (null).
        static int? CompareValues(ValueEval x, ValueEval y)
        {
            if (x is NumberEval nx && y is NumberEval ny) return nx.NumberValue.CompareTo(ny.NumberValue);
            if (x is StringEval sx && y is StringEval sy) return string.Compare(sx.StringValue, sy.StringValue, StringComparison.OrdinalIgnoreCase);
            if (x is BoolEval bx && y is BoolEval by) return bx.BooleanValue.CompareTo(by.BooleanValue);
            return null;
        }

        // DEC2BIN/DEC2OCT/DEC2HEX: negative numbers are ten-digit two's complement and ignore `places`.
        static ValueEval DecTo(Ctx c, ValueEval[] a, int radix, int digits)
        {
            Need(a, 1, 2);
            var n = Math.Truncate(c.Num(a[0]));
            var bits = radix == 2 ? 10 : radix == 8 ? 30 : 40;
            var limit = Math.Pow(2, bits - 1);
            if (n < -limit || n >= limit) throw XlError.Num;
            if (n < 0) return S(ToBase((long)(Math.Pow(2, bits) + n), radix));
            var text = ToBase((long)n, radix);
            if (a.Length > 1 && !Ctx.IsMissing(a[1]))
            {
                var places = (int)Math.Truncate(c.Num(a[1]));
                if (places <= 0 || places > 10 || places < text.Length) throw XlError.Num;
                text = text.PadLeft(places, '0');
            }
            return S(text);
        }

        /// <summary>Excel wildcards: ? is one character, * any run, ~ escapes the next character.</summary>
        static string WildcardToRegex(string pattern)
        {
            var sb = new StringBuilder();
            for (var i = 0; i < pattern.Length; i++)
            {
                var ch = pattern[i];
                if (ch == '~' && i + 1 < pattern.Length) sb.Append(Regex.Escape(pattern[++i].ToString()));
                else if (ch == '?') sb.Append('.');
                else if (ch == '*') sb.Append(".*?");
                else sb.Append(Regex.Escape(ch.ToString()));
            }
            return sb.ToString();
        }

        static void RejectBooleans(ValueEval[] a)
        {
            if (a.Any(x => x is BoolEval)) throw XlError.Value;
        }

        static ValueEval WorkdayIntl(Ctx c, ValueEval[] a)
        {
            Need(a, 2, 4);
            if (a.Take(2).Any(x => x is BoolEval)) throw XlError.Value;
            var start = Math.Floor(c.Num(a[0])); var days = (int)Math.Truncate(c.Num(a[1]));
            var weekend = Weekend(c, a, 2); var holidays = Holidays(c, a, 3);
            var step = days >= 0 ? 1 : -1; var left = Math.Abs(days); var d = start;
            Dow(start);
            while (left > 0)
            {
                d += step;
                if (!weekend[Dow(d)] && !holidays.Contains(d)) left--;
            }
            return N(d);
        }

        // Puts the corrections and the extra names where NPOI will find them. Called once.
        static void InstallIntoNpoi()
        {
            DefineCorrections();

            // 1. Analysis ToolPak names NPOI lists as not supported: plain registration.
            var missingAtp = new HashSet<string>(AnalysisToolPak.GetNotSupportedFunctionNames(), StringComparer.OrdinalIgnoreCase);
            foreach (var kv in Map.Concat(Corrections.Where(k => !Map.ContainsKey(k.Key))))
            {
                if (!missingAtp.Contains(kv.Key)) continue;
                try { AnalysisToolPak.RegisterFunction(kv.Key, kv.Value); } catch (Exception) { /* leave it as NPOI has it */ }
            }

            // 2. Corrections: wrap whatever NPOI has, so it only takes over inside an OpenSheet evaluation.
            Array builtins = null;
            FieldInfo builtinField = null;
            try
            {
                builtinField = typeof(FunctionEval).GetField("functions", BindingFlags.NonPublic | BindingFlags.Static);
                builtins = builtinField?.GetValue(null) as Array;
            }
            catch (Exception) { /* different NPOI layout: skip the corrections, keep everything else */ }

            foreach (var name in Guarded)
            {
                try
                {
                    if (AnalysisToolPak.IsATPFunction(name))
                    {
                        var atpField = typeof(AnalysisToolPak).GetField("_functionsByName", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);
                        var atpTable = atpField?.GetValue(atpField.IsStatic ? null : AnalysisToolPak.instance) as Dictionary<string, FreeRefFunction>;
                        if (atpTable != null && atpTable.TryGetValue(name.ToUpperInvariant(), out var atpCurrent)) atpTable[name.ToUpperInvariant()] = new GuardedAtp(atpCurrent, RejectsBooleans(name));
                        continue;
                    }
                    var meta = NPOI.SS.Formula.Function.FunctionMetadataRegistry.GetFunctionByName(name);
                    if (meta == null || builtins == null) continue;
                    if (builtins.GetValue(meta.Index) is NPOI.SS.Formula.Functions.Function original) builtins.SetValue(new GuardedBuiltin(original, RejectsBooleans(name)), meta.Index);
                }
                catch (Exception) { /* a guard that cannot be installed leaves NPOI's own behavior */ }
            }

            foreach (var kv in Corrections)
            {
                try
                {
                    if (AnalysisToolPak.IsATPFunction(kv.Key) && !missingAtp.Contains(kv.Key))
                    {
                        // The public API refuses to replace a working function, so swap the entry in the private table.
                        var field = typeof(AnalysisToolPak).GetField("_functionsByName", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);
                        var table = field?.GetValue(field.IsStatic ? null : AnalysisToolPak.instance) as Dictionary<string, FreeRefFunction>;
                        if (table != null && table.TryGetValue(kv.Key.ToUpperInvariant(), out var current))
                            table[kv.Key.ToUpperInvariant()] = new ScopedAtp(current, kv.Value);
                        continue;
                    }
                    var meta = NPOI.SS.Formula.Function.FunctionMetadataRegistry.GetFunctionByName(kv.Key);
                    if (meta == null || builtins == null) continue;
                    var original = builtins.GetValue(meta.Index) as NPOI.SS.Formula.Functions.Function;
                    if (original == null) continue;
                    builtins.SetValue(new ScopedBuiltin(original, kv.Value), meta.Index);
                }
                catch (Exception) { /* a correction that cannot be installed leaves NPOI's own behavior */ }
            }
        }
    }
}
