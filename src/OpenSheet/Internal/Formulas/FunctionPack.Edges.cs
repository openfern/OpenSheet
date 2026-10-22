using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;

namespace OpenSheet.Internal.Formulas
{
    /// <summary>NPOI function that, while OpenSheet evaluates, hands back the first error among its arguments before doing anything else.</summary>
    sealed class GuardedBuiltin : NPOI.SS.Formula.Functions.Function
    {
        readonly NPOI.SS.Formula.Functions.Function _original;
        readonly bool _noBooleans;
        public GuardedBuiltin(NPOI.SS.Formula.Functions.Function original, bool noBooleans) { _original = original; _noBooleans = noBooleans; }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) =>
            (EvaluationScope.Active ? FunctionPack.Guard(args, _noBooleans) : null) ?? _original.Evaluate(args, srcRowIndex, srcColumnIndex);
    }

    sealed class GuardedAtp : FreeRefFunction
    {
        readonly FreeRefFunction _original;
        readonly bool _noBooleans;
        public GuardedAtp(FreeRefFunction original, bool noBooleans) { _original = original; _noBooleans = noBooleans; }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) =>
            (EvaluationScope.Active ? FunctionPack.Guard(args, _noBooleans) : null) ?? _original.Evaluate(args, ec);
    }

    static partial class FunctionPack
    {
        /// <summary>Functions NPOI computes correctly but whose answer to an error argument is #VALUE! where Excel passes the error on.</summary>
        static readonly string[] Guarded = { "EDATE", "EOMONTH", "WEEKNUM", "PERCENTILE", "QUOTIENT", "MROUND", "ISODD", "ISEVEN", "YEARFRAC", "FACTDOUBLE", "COMPLEX" };

        /// <summary>The ones among them that, like the rest of the Analysis ToolPak, answer #VALUE! to TRUE and FALSE.</summary>
        static readonly HashSet<string> NoBooleanNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "EDATE", "EOMONTH", "WEEKNUM", "QUOTIENT", "MROUND", "ISODD", "ISEVEN", "YEARFRAC", "FACTDOUBLE", "COMPLEX" };

        internal static bool RejectsBooleans(string name) => NoBooleanNames.Contains(name);

        /// <summary>What a guarded function answers before looking at anything else: the first error among its arguments, or #VALUE! for a boolean; null otherwise.</summary>
        internal static ErrorEval Guard(ValueEval[] args, bool noBooleans)
        {
            var error = FirstError(args);
            if (error != null) return error;
            return noBooleans && args.Any(x => x is BoolEval) ? ErrorEval.VALUE_INVALID : null;
        }

        /// <summary>The first error among the arguments (typed in or held by a single cell), or null.</summary>
        internal static ErrorEval FirstError(ValueEval[] args)
        {
            foreach (var a in args)
            {
                if (a is ErrorEval e) return e;
                if (a is RefEval r && r.GetInnerValueEval(r.FirstSheetIndex) is ErrorEval held) return held;
            }
            return null;
        }

        // ---- Excel's weekday is arithmetic on the serial; AND / OR / COUNT read typed-in values differently from NPOI ----

        static void DefineEdgeCorrections()
        {
            Correct("WEEKDAY", (c, a) =>
            {
                Need(a, 1, 2);
                var serial = Math.Floor(c.Num(a[0])); var type = c.Int(a, 1, 1);
                var sunday0 = Dow(serial);   // 0 = Sunday
                switch (type)
                {
                    case 1: return N(sunday0 + 1);
                    case 2: return N((sunday0 + 6) % 7 + 1);
                    case 3: return N((sunday0 + 6) % 7);
                    default:
                        if (type < 11 || type > 17) throw XlError.Num;
                        return N(((sunday0 - (type - 10) % 7) % 7 + 7) % 7 + 1);
                }
            });

            Correct("COUNT", (c, a) =>
            {
                var n = 0;
                foreach (var arg in a)
                {
                    if (arg is AreaEval) { n += c.Area(arg).Flat().Count(x => x is NumberEval); continue; }
                    if (arg is RefEval r) { if (r.GetInnerValueEval(r.FirstSheetIndex) is NumberEval) n++; continue; }
                    // typed straight into the formula: numbers, booleans and text that reads as a number all count
                    if (arg is NumberEval || arg is BoolEval || (arg is StringEval s && Ctx.ParseNumber(s.StringValue) != null)) n++;
                }
                return N(n);
            });

            Correct("AND", (c, a) => Logical(c, a, all: true));
            Correct("OR", (c, a) => Logical(c, a, all: false));

            Correct("IFS", (c, a) =>
            {
                if (a.Length == 0 || a.Length % 2 != 0) throw XlError.Value;
                for (var i = 0; i < a.Length; i += 2)
                {
                    bool condition;
                    switch (c.Scalar(a[i]))
                    {
                        case BoolEval b: condition = b.BooleanValue; break;
                        case NumberEval n: condition = n.NumberValue != 0; break;
                        case BlankEval _: condition = false; break;
                        case ErrorEval e: throw new XlError(e);
                        default: throw XlError.Value;   // text is not a condition
                    }
                    if (condition) return c.Scalar(a[i + 1]);
                }
                throw XlError.NA;
            });

            Correct("RANK", (c, a) => Rank(c, a, average: false));   // RANK is RANK.EQ under its older name
            // a regression through fewer than two points, or through points that do not vary in x, is #DIV/0!
            Corrections["FORECAST"] = Map["FORECAST.LINEAR"];
            Correct("SLOPE", (c, a) =>
            {
                Need(a, 2, 2); Pairs(c, a[0], a[1], out var y, out var x);
                if (x.Count < 2 || Cov(x, x, true) == 0) throw XlError.Div0;
                return N(Cov(x, y, true) / Cov(x, x, true));
            });
            Correct("INTERCEPT", (c, a) =>
            {
                Need(a, 2, 2); Pairs(c, a[0], a[1], out var y, out var x);
                if (x.Count < 2 || Cov(x, x, true) == 0) throw XlError.Div0;
                return N(y.Average() - Cov(x, y, true) / Cov(x, x, true) * x.Average());
            });
            Correct("LOG", (c, a) =>
            {
                Need(a, 1, 2);
                double x = c.Num(a[0]), b = c.Num(a, 1, 10);
                if (x <= 0 || b <= 0) throw XlError.Num;
                if (b == 1) throw XlError.Div0;
                return N(Math.Log(x) / Math.Log(b));
            });

            // spread of a sample / population: one value is not enough for a sample, none for a population (#DIV/0!, not #NUM!)
            foreach (var name in new[] { "VAR", "VAR.S" }) Correct(name, (c, a) => Dispersion(c, a, sample: true, root: false));
            foreach (var name in new[] { "VARP", "VAR.P" }) Correct(name, (c, a) => Dispersion(c, a, sample: false, root: false));
            foreach (var name in new[] { "STDEV", "STDEV.S" }) Correct(name, (c, a) => Dispersion(c, a, sample: true, root: true));
            foreach (var name in new[] { "STDEVP", "STDEV.P" }) Correct(name, (c, a) => Dispersion(c, a, sample: false, root: true));

            // NPOI wants a second argument that Excel makes optional (0)
            Correct("DELTA", (c, a) => { Need(a, 1, 2); if (a.Any(x => x is BoolEval)) throw XlError.Value; return N(c.Num(a[0]) == c.Num(a, 1, 0) ? 1 : 0); });
            Correct("STANDARDIZE", (c, a) =>
            {
                Need(a, 3, 3);
                double x = c.Num(a[0]), mean = c.Num(a[1]), sd = c.Num(a[2]);
                if (sd <= 0) throw XlError.Num;
                return N((x - mean) / sd);
            });

            // anything that is not a number counts as 0 in the products (TRUE, text, empty cells); an error anywhere is the answer
            Correct("SUMPRODUCT", (c, a) =>
            {
                if (a.Length == 0) throw XlError.Value;
                var grids = a.Select(x => c.Area(x)).ToList();
                if (grids.Any(g => g.Rows != grids[0].Rows || g.Columns != grids[0].Columns)) throw XlError.Value;
                double total = 0;
                for (var r = 0; r < grids[0].Rows; r++)
                    for (var col = 0; col < grids[0].Columns; col++)
                    {
                        var product = 1.0;
                        foreach (var g in grids)
                        {
                            var v = g.Cells[r, col];
                            if (v is ErrorEval e) throw new XlError(e);
                            product *= v is NumberEval n ? n.NumberValue : 0;
                        }
                        total += product;
                    }
                return N(total);
            });

            // the "A" versions count text as 0 and TRUE as 1, but leave empty cells out
            Correct("MINA", (c, a) => { var xs = c.NumbersWithText(a); return N(xs.Count == 0 ? 0 : xs.Min()); });
            Correct("MAXA", (c, a) => { var xs = c.NumbersWithText(a); return N(xs.Count == 0 ? 0 : xs.Max()); });

            // WORKDAY and NETWORKDAYS are the .INTL functions with the usual Saturday-Sunday weekend (and, like the rest of the Analysis ToolPak, no booleans)
            Correct("WORKDAY", (c, a) =>
            {
                Need(a, 2, 3);
                if (a.Take(2).Any(x => x is BoolEval)) throw XlError.Value;
                return WorkdayIntl(c, new[] { a[0], a[1], MissingArgEval.instance, a.Length > 2 ? a[2] : MissingArgEval.instance });
            });
            Correct("NETWORKDAYS", (c, a) =>
            {
                Need(a, 2, 3);
                if (a.Take(2).Any(x => x is BoolEval)) throw XlError.Value;
                return NetworkdaysIntl(c, new[] { a[0], a[1], MissingArgEval.instance, a.Length > 2 ? a[2] : MissingArgEval.instance });
            });

            Register("PERCENTRANK", inclusive: true); Register("PERCENTRANK.INC", inclusive: true); Register("PERCENTRANK.EXC", inclusive: false);

            void Register(string name, bool inclusive)
            {
                Add(name, (c, a) => PercentRank(c, a, inclusive));
                Corrections[name] = Map[name];
            }
        }

        static ValueEval Dispersion(Ctx c, ValueEval[] a, bool sample, bool root)
        {
            var xs = c.Numbers(a);
            if (xs.Count < (sample ? 2 : 1)) throw XlError.Div0;
            var v = Variance(xs, sample);
            return N(root ? Math.Sqrt(v) : v);
        }

        // AND / OR: booleans and numbers decide; text and empty cells are left out; an error stops everything; nothing logical at all is #VALUE!
        static ValueEval Logical(Ctx c, ValueEval[] a, bool all)
        {
            var any = false; var result = all;
            void Take(ValueEval v)
            {
                if (v is ErrorEval e) throw new XlError(e);
                bool value;
                if (v is BoolEval b) value = b.BooleanValue;
                else if (v is NumberEval n) value = n.NumberValue != 0;
                else return;   // text, even "TRUE", and empty cells are left out wherever they come from
                any = true;
                result = all ? result && value : result || value;
            }
            foreach (var arg in a)
            {
                if (arg is AreaEval) foreach (var x in c.Area(arg).Flat()) Take(x);
                else if (arg is RefEval r) Take(r.GetInnerValueEval(r.FirstSheetIndex));
                else Take(arg);
            }
            if (!any) throw XlError.Value;
            return B(result);
        }

        // PERCENTRANK(array, x, [significance]): where x sits among the values, 0 to 1, cut (not rounded) to the digits asked for
        static ValueEval PercentRank(Ctx c, ValueEval[] a, bool inclusive)
        {
            Need(a, 2, 3);
            var xs = c.Numbers(new[] { a[0] }); xs.Sort();
            var x = c.Num(a[1]); var digits = c.Int(a, 2, 3);
            if (digits < 1) throw XlError.Num;
            if (xs.Count == 0) throw XlError.NA;
            if (x < xs[0] || x > xs[xs.Count - 1]) throw XlError.NA;
            double rank;
            var first = xs.FindIndex(v => v >= x);
            if (xs[first] == x)
            {
                if (inclusive) rank = xs.Count == 1 ? 1 : (double)first / (xs.Count - 1);
                else rank = (first + 1.0) / (xs.Count + 1);
            }
            else
            {
                var lo = first - 1; var fraction = (x - xs[lo]) / (xs[first] - xs[lo]);
                rank = inclusive ? (lo + fraction) / (xs.Count - 1) : (lo + 1 + fraction) / (xs.Count + 1);
            }
            var scale = Math.Pow(10, digits);
            return N(Math.Floor(rank * scale + 1e-9) / scale);
        }
    }
}
