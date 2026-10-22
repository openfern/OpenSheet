using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.Formula.Eval;

namespace OpenSheet.Internal.Formulas
{
    static partial class FunctionPack
    {
        // ---- small helpers shared by all groups ----

        static ValueEval N(double d) => double.IsNaN(d) || double.IsInfinity(d) ? ErrorEval.NUM_ERROR : (ValueEval)new NumberEval(d);
        static ValueEval S(string s) => new StringEval(s);
        static ValueEval B(bool b) => BoolEval.ValueOf(b);

        static void Need(ValueEval[] a, int min, int max)
        {
            if (a.Length < min || a.Length > max) throw XlError.Value;
        }

        static void Num1(string name, Func<double, double> f) => Add(name, (c, a) => { Need(a, 1, 1); return N(f(c.Num(a[0]))); });
        static void Num2(string name, Func<double, double, double> f) => Add(name, (c, a) => { Need(a, 2, 2); return N(f(c.Num(a[0]), c.Num(a[1]))); });
        static void Num3(string name, Func<double, double, double, double> f) => Add(name, (c, a) => { Need(a, 3, 3); return N(f(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]))); });

        static void Define()
        {
            DefineMath();
            DefineStats();
            DefineDistributions();
            DefineText();
            DefineInfo();
            DefineDates();
            DefineFinance();
            DefineEngineering();
            // Analysis ToolPak functions answer #VALUE! to TRUE and FALSE where Excel's other math functions read 1 and 0
            NoBooleans(int.MaxValue, "EFFECT", "ERF", "ERF.PRECISE", "ERFC", "ERFC.PRECISE", "FVSCHEDULE", "GCD", "GESTEP", "LCM", "MULTINOMIAL", "NOMINAL", "SQRTPI", "XNPV", "CONVERT", "SERIESSUM");
            NoBooleans(2, "NETWORKDAYS.INTL");   // its third argument, the weekend, may be TRUE
        }

        static void NoBooleans(int firstArguments, params string[] names)
        {
            foreach (var name in names)
            {
                var inner = Map[name].Impl;
                Map[name] = new PackFunction((c, a) =>
                {
                    if (a.Take(firstArguments).Any(x => x is BoolEval)) throw XlError.Value;
                    return inner(c, a);
                });
            }
        }

        static double Fact(double n)
        {
            n = Math.Floor(n);
            if (n < 0 || n > 170) throw XlError.Num;
            double r = 1;
            for (var i = 2; i <= n; i++) r *= i;
            return r;
        }

        static double Gcd(double a, double b)
        {
            while (b != 0) { var t = a % b; a = b; b = t; }
            return a;
        }

        static void DefineMath()
        {
            Add("GCD", (c, a) =>
            {
                var xs = c.Numbers(a);
                if (xs.Count == 0) throw XlError.Value;
                if (xs.Any(x => x < 0 || x >= 9007199254740992.0)) throw XlError.Num;
                return N(xs.Select(Math.Floor).Aggregate(0.0, Gcd));
            });
            Add("LCM", (c, a) =>
            {
                var xs = c.Numbers(a);
                if (xs.Count == 0) throw XlError.Value;
                if (xs.Any(x => x < 0 || x >= 9007199254740992.0)) throw XlError.Num;
                double l = 1;
                foreach (var x in xs.Select(Math.Floor))
                {
                    if (x == 0) return N(0);
                    l = l / Gcd(l, x) * x;
                }
                return N(l);
            });
            Add("MULTINOMIAL", (c, a) =>
            {
                var xs = c.Numbers(a).Select(Math.Floor).ToList();
                if (xs.Any(x => x < 0)) throw XlError.Num;
                var v = Fact(xs.Sum());
                foreach (var x in xs) v /= Fact(x);
                return N(Math.Round(v));
            });
            Num1("SQRTPI", x => x < 0 ? throw XlError.Num : Math.Sqrt(x * Math.PI));
            Num1("SEC", x => Trig(x) ? 1 / Math.Cos(x) : 0);
            Num1("CSC", x => { Trig(x); var s = Math.Sin(x); return s == 0 ? throw XlError.Div0 : 1 / s; });
            Num1("COT", x => { Trig(x); var t = Math.Tan(x); return t == 0 ? throw XlError.Div0 : 1 / t; });
            Num1("SECH", x => 1 / Math.Cosh(x));
            Num1("CSCH", x => x == 0 ? throw XlError.Div0 : 1 / Math.Sinh(x));
            Num1("COTH", x => x == 0 ? throw XlError.Div0 : 1 / Math.Tanh(x));
            Num1("ACOT", x => Math.PI / 2 - Math.Atan(x));
            Num1("ACOTH", x => Math.Abs(x) <= 1 ? throw XlError.Num : 0.5 * Math.Log((x + 1) / (x - 1)));
            Num2("COMBINA", (n, k) =>
            {
                n = Math.Floor(n); k = Math.Floor(k);
                if (n < 0 || k < 0 || (n == 0 && k > 0)) throw XlError.Num;
                if (n == 0 && k == 0) return 1;
                return Math.Round(Math.Exp(Special.LogChoose(n + k - 1, k)));
            });
            Num2("PERMUT", (n, k) =>
            {
                n = Math.Floor(n); k = Math.Floor(k);
                if (n <= 0 || k < 0 || k > n) throw XlError.Num;
                return Math.Round(Math.Exp(Special.LogGamma(n + 1) - Special.LogGamma(n - k + 1)));
            });
            Num2("PERMUTATIONA", (n, k) =>
            {
                n = Math.Floor(n); k = Math.Floor(k);
                if (n < 0 || k < 0) throw XlError.Num;
                return Math.Pow(n, k);
            });
            Add("CEILING.PRECISE", (c, a) => N(Round(c, a, up: true)));
            Add("ISO.CEILING", (c, a) => N(Round(c, a, up: true)));
            Add("FLOOR.PRECISE", (c, a) => N(Round(c, a, up: false)));
            Add("CEILING.MATH", (c, a) => N(RoundMath(c, a, ceiling: true)));
            Add("FLOOR.MATH", (c, a) => N(RoundMath(c, a, ceiling: false)));
            Add("ARABIC", (c, a) => { Need(a, 1, 1); return N(Arabic(c.Str(a[0]))); });
            Add("ROMAN", (c, a) =>
            {
                Need(a, 1, 2);
                var n = Math.Floor(c.Num(a[0]));
                var form = a.Length > 1 ? (int)Math.Truncate(c.Num(a[1])) : 0;
                if (n < 0 || n > 3999 || form < 0 || form > 4) throw XlError.Value;
                return S(Roman((int)n, form));
            });
            Add("BASE", (c, a) =>
            {
                Need(a, 2, 3);
                var n = Math.Floor(c.Num(a[0])); var radix = c.Int(a[1]); var min = c.Int(a, 2, 0);
                if (n < 0 || n >= 9007199254740992.0 || radix < 2 || radix > 36 || min < 0 || min > 255) throw XlError.Num;
                return S(ToBase((long)n, radix).PadLeft(min, '0'));
            });
            Add("DECIMAL", (c, a) =>
            {
                Need(a, 2, 2);
                var text = c.Str(a[0]).Trim(); var radix = c.Int(a[1]);
                if (radix < 2 || radix > 36) throw XlError.Num;
                double v = 0;
                foreach (var ch in text.ToUpperInvariant())
                {
                    var d = char.IsDigit(ch) ? ch - '0' : ch >= 'A' && ch <= 'Z' ? ch - 'A' + 10 : 99;
                    if (d >= radix) throw XlError.Num;
                    v = v * radix + d;
                }
                return N(v);
            });
            Add("SERIESSUM", (c, a) =>
            {
                Need(a, 4, 4);
                double x = c.Num(a[0]), n = c.Num(a[1]), m = c.Num(a[2]);
                var coef = Series(c, a[3]).ToList();
                double sum = 0;
                for (var i = 0; i < coef.Count; i++) sum += coef[i] * Math.Pow(x, n + i * m);
                return N(sum);
            });
            Add("AGGREGATE", Aggregate);
        }

        // Excel refuses the angle of these trigonometric functions when it is 2^27 or more
        static bool Trig(double x) => Math.Abs(x) < 134217728 ? true : throw XlError.Num;

        static double Round(Ctx c, ValueEval[] a, bool up)
        {
            Need(a, 1, 2);
            var x = c.Num(a[0]); var s = Math.Abs(c.Num(a, 1, 1));
            if (s == 0) return 0;
            return (up ? Math.Ceiling(x / s) : Math.Floor(x / s)) * s;
        }

        static double RoundMath(Ctx c, ValueEval[] a, bool ceiling)
        {
            Need(a, 1, 3);
            var x = c.Num(a[0]); var s = Math.Abs(c.Num(a, 1, 1)); var mode = c.Num(a, 2, 0) != 0;
            if (s == 0) return 0;
            if (ceiling) return x >= 0 || !mode ? Math.Ceiling(x / s) * s : -Math.Ceiling(-x / s) * s;
            return x >= 0 || !mode ? Math.Floor(x / s) * s : -Math.Floor(-x / s) * s;
        }

        static string ToBase(long n, int radix)
        {
            if (n == 0) return "0";
            var sb = new StringBuilder();
            while (n > 0) { var d = (int)(n % radix); sb.Insert(0, (char)(d < 10 ? '0' + d : 'A' + d - 10)); n /= radix; }
            return sb.ToString();
        }

        static double Arabic(string text)
        {
            text = text.Trim().ToUpperInvariant();
            if (text.Length == 0) return 0;
            var neg = text.StartsWith("-", StringComparison.Ordinal);
            if (neg) text = text.Substring(1);
            int Val(char ch) { switch (ch) { case 'I': return 1; case 'V': return 5; case 'X': return 10; case 'L': return 50; case 'C': return 100; case 'D': return 500; case 'M': return 1000; default: throw XlError.Value; } }
            double total = 0;
            for (var i = 0; i < text.Length; i++)
            {
                var v = Val(text[i]);
                total += i + 1 < text.Length && v < Val(text[i + 1]) ? -v : v;
            }
            return neg ? -total : total;
        }

        // Form 0 is the classic notation; forms 1-4 shorten it step by step (499 = CDXCIX, LDVLIV, XDIX, VDIV, ID).
        static string Roman(int n, int form)
        {
            if (n == 0) return "";
            var chars = new[] { 'M', 'D', 'C', 'L', 'X', 'V', 'I' };
            var values = new[] { 1000, 500, 100, 50, 10, 5, 1 };
            var sb = new StringBuilder();
            for (var i = 0; i <= 3; i++)
            {
                var idx = 2 * i;
                var digit = n / values[idx];
                if (digit % 5 == 4)
                {
                    var idx2 = digit == 4 ? idx - 1 : idx - 2;
                    var steps = 0;
                    while (steps < form && idx < 6)
                    {
                        steps++;
                        if (values[idx2] - values[idx + 1] <= n) idx++;
                        else steps = form;
                    }
                    sb.Append(chars[idx]).Append(chars[idx2]);
                    n += values[idx];
                    n -= values[idx2];
                }
                else
                {
                    if (digit > 4) sb.Append(chars[idx - 1]);
                    digit %= 5;
                    while (digit-- > 0) sb.Append(chars[idx]);
                    n %= values[idx];
                }
            }
            return sb.ToString();
        }

        // ---- AGGREGATE ----

        static ValueEval Aggregate(Ctx c, ValueEval[] a)
        {
            if (a.Length < 3) throw XlError.Value;
            var fn = c.Int(a[0]); var opt = c.Int(a[1]);
            if (fn < 1 || fn > 19 || opt < 0 || opt > 7) throw XlError.Value;
            if (fn <= 13)   // the first thirteen work on references, not on typed-in numbers
                foreach (var r in a.Skip(2))
                {
                    if (r is ErrorEval err) throw new XlError(err);
                    if (!(r is AreaEval) && !(r is RefEval)) throw XlError.Value;
                }
            var skipErrors = opt == 2 || opt == 3 || opt == 6 || opt == 7;
            List<double> Collect(IEnumerable<ValueEval> refs)
            {
                var list = new List<double>();
                foreach (var r in refs)
                    foreach (var x in r is AreaEval ? c.Area(r).Flat() : new[] { c.Scalar(r) })
                    {
                        if (x is NumberEval n) list.Add(n.NumberValue);
                        else if (x is ErrorEval e) { if (!skipErrors) throw new XlError(e); }
                    }
                return list;
            }

            if (fn >= 14)
            {
                if (a.Length < 4) throw XlError.Value;
                var data = Collect(new[] { a[2] }); var k = c.Num(a[3]);
                data.Sort();
                switch (fn)
                {
                    case 14: { var i = (int)Math.Ceiling(k); if (i < 1 || i > data.Count) throw XlError.Num; return N(data[data.Count - i]); }
                    case 15: { var i = (int)Math.Ceiling(k); if (i < 1 || i > data.Count) throw XlError.Num; return N(data[i - 1]); }
                    case 16: return N(PercentileInc(data, k));
                    case 17: return N(PercentileInc(data, Math.Truncate(k) / 4));
                    case 18: return N(PercentileExc(data, k));
                    default: return N(PercentileExc(data, Math.Truncate(k) / 4));
                }
            }
            var xs = Collect(a.Skip(2));
            switch (fn)
            {
                case 1: if (xs.Count == 0) throw XlError.Div0; return N(xs.Average());
                case 2: return N(xs.Count);
                case 3: return N(a.Skip(2).Sum(r => (r is AreaEval ? c.Area(r).Flat() : new[] { c.Scalar(r) }).Count(x => !(x is BlankEval) && !(x is MissingArgEval) && (!skipErrors || !(x is ErrorEval)))));
                case 4: if (xs.Count == 0) return N(0); return N(xs.Max());
                case 5: if (xs.Count == 0) return N(0); return N(xs.Min());
                case 6: return N(xs.Aggregate(1.0, (p, x) => p * x));
                case 7: if (xs.Count < 2) throw XlError.Div0; return N(Math.Sqrt(Variance(xs, sample: true)));
                case 8: if (xs.Count < 1) throw XlError.Div0; return N(Math.Sqrt(Variance(xs, sample: false)));
                case 9: return N(xs.Sum());
                case 10: if (xs.Count < 2) throw XlError.Div0; return N(Variance(xs, sample: true));
                case 11: if (xs.Count < 1) throw XlError.Div0; return N(Variance(xs, sample: false));
                case 12: if (xs.Count == 0) throw XlError.Num; return N(Median(xs));
                default: return N(ModeOf(xs));
            }
        }
    }
}
