using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Formula.Eval;

namespace OpenSheet.Internal.Formulas
{
    static partial class FunctionPack
    {
        static double Variance(IList<double> xs, bool sample)
        {
            var mean = xs.Average();
            var ss = xs.Sum(x => (x - mean) * (x - mean));
            return ss / (sample ? xs.Count - 1 : xs.Count);
        }

        static double Median(List<double> xs)
        {
            var s = xs.OrderBy(x => x).ToList();
            var n = s.Count;
            return n % 2 == 1 ? s[n / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
        }

        static double ModeOf(List<double> xs)
        {
            var groups = xs.GroupBy(x => x).ToList();
            var max = groups.Count == 0 ? 0 : groups.Max(g => g.Count());
            if (max < 2) throw XlError.NA;
            var top = new HashSet<double>(groups.Where(g => g.Count() == max).Select(g => g.Key));
            return xs.First(top.Contains); // ties go to the value that appears first
        }

        static double PercentileInc(List<double> sorted, double k)
        {
            if (sorted.Count == 0 || k < 0 || k > 1) throw XlError.Num;
            sorted.Sort();
            var rank = k * (sorted.Count - 1);
            var lo = (int)Math.Floor(rank); var frac = rank - lo;
            return lo + 1 < sorted.Count ? sorted[lo] + frac * (sorted[lo + 1] - sorted[lo]) : sorted[lo];
        }

        static double PercentileExc(List<double> sorted, double k)
        {
            sorted.Sort();
            var n = sorted.Count;
            if (n == 0 || k <= 0 || k >= 1) throw XlError.Num;
            var rank = k * (n + 1);
            if (rank < 1 || rank > n) throw XlError.Num;
            var lo = (int)Math.Floor(rank); var frac = rank - lo;
            return lo < n ? sorted[lo - 1] + frac * (sorted[lo] - sorted[lo - 1]) : sorted[lo - 1];
        }

        // Pairs of numbers from two equally shaped ranges; a pair is dropped when either side is not a number.
        static void Pairs(Ctx c, ValueEval y, ValueEval x, out List<double> ys, out List<double> xs)
        {
            var gy = c.Area(y).Flat().ToList(); var gx = c.Area(x).Flat().ToList();
            if (gy.Count != gx.Count) throw XlError.NA;
            ys = new List<double>(); xs = new List<double>();
            for (var i = 0; i < gy.Count; i++)
            {
                if (gy[i] is ErrorEval ey) throw new XlError(ey);
                if (gx[i] is ErrorEval ex) throw new XlError(ex);
                if (gy[i] is NumberEval ny && gx[i] is NumberEval nx) { ys.Add(ny.NumberValue); xs.Add(nx.NumberValue); }
            }
        }

        static double Cov(List<double> a, List<double> b, bool sample)
        {
            var ma = a.Average(); var mb = b.Average();
            double s = 0;
            for (var i = 0; i < a.Count; i++) s += (a[i] - ma) * (b[i] - mb);
            return s / (sample ? a.Count - 1 : a.Count);
        }

        static void DefineStats()
        {
            Add("AVERAGEA", (c, a) =>
            {
                var xs = c.NumbersWithText(a);
                if (xs.Count == 0) throw XlError.Div0;
                return N(xs.Average());
            });
            Impl modeSngl = (c, a) =>
            {
                // a lone reference to an empty cell is an error here, while text typed into the formula is simply left out
                if (a.Any(x => x is RefEval && c.Scalar(x) is BlankEval)) throw XlError.Value;
                return N(ModeOf(c.Numbers(a.Where(x => !(x is StringEval)).ToArray())));
            };   // text typed into the formula is ignored, even "2"
            Add("MODE.SNGL", modeSngl);
            Add("RANK.EQ", (c, a) => Rank(c, a, average: false));
            Add("RANK.AVG", (c, a) => Rank(c, a, average: true));
            Add("PERCENTILE.INC", (c, a) => { Need(a, 2, 2); var xs = c.Numbers(new[] { a[0] }); return N(PercentileInc(xs, c.Num(a[1]))); });
            Add("PERCENTILE.EXC", (c, a) => { Need(a, 2, 2); var xs = c.Numbers(new[] { a[0] }); return N(PercentileExc(xs, c.Num(a[1]))); });
            Impl quartile = (c, a) =>
            {
                Need(a, 2, 2);
                var xs = c.Numbers(new[] { a[0] }); var q = (int)Math.Truncate(c.Num(a[1]));
                if (q < 0 || q > 4) throw XlError.Num;
                return N(PercentileInc(xs, q / 4.0));
            };
            Add("QUARTILE", quartile);
            Add("QUARTILE.INC", quartile);
            Add("QUARTILE.EXC", (c, a) =>
            {
                Need(a, 2, 2);
                var xs = c.Numbers(new[] { a[0] }); var q = (int)Math.Truncate(c.Num(a[1]));
                if (q < 1 || q > 3) throw XlError.Num;
                return N(PercentileExc(xs, q / 4.0));
            });
            Add("GEOMEAN", (c, a) =>
            {
                var xs = c.Numbers(a);
                if (xs.Count == 0 || xs.Any(x => x <= 0)) throw XlError.Num;
                return N(Math.Exp(xs.Average(Math.Log)));
            });
            Add("HARMEAN", (c, a) =>
            {
                var xs = c.Numbers(a);
                if (xs.Count == 0) throw XlError.NA;
                if (xs.Any(x => x <= 0)) throw XlError.Num;
                return N(xs.Count / xs.Sum(x => 1 / x));
            });
            Add("TRIMMEAN", (c, a) =>
            {
                Need(a, 2, 2);
                var xs = c.Numbers(new[] { a[0] }); var p = c.Num(a[1]);
                if (xs.Count == 0 || p < 0 || p >= 1) throw XlError.Num;
                var cut = (int)Math.Floor(xs.Count * p / 2);
                var kept = xs.OrderBy(x => x).Skip(cut).Take(xs.Count - 2 * cut).ToList();
                return N(kept.Average());
            });
            Add("CORREL", (c, a) =>
            {
                Need(a, 2, 2); Pairs(c, a[0], a[1], out var y, out var x);
                if (x.Count < 2) throw XlError.Div0;
                var d = Math.Sqrt(Cov(x, x, true) * Cov(y, y, true));
                if (d == 0) throw XlError.Div0;
                return N(Cov(x, y, true) / d);
            });
            Impl covP = (c, a) => { Need(a, 2, 2); Pairs(c, a[0], a[1], out var y, out var x); if (x.Count == 0) throw XlError.Div0; return N(Cov(x, y, false)); };
            Add("COVAR", covP);
            Add("COVARIANCE.P", covP);
            Add("COVARIANCE.S", (c, a) => { Need(a, 2, 2); Pairs(c, a[0], a[1], out var y, out var x); if (x.Count < 2) throw XlError.Div0; return N(Cov(x, y, true)); });
            Add("RSQ", (c, a) =>
            {
                Need(a, 2, 2); Pairs(c, a[0], a[1], out var y, out var x);
                if (x.Count < 2) throw XlError.Div0;
                var d = Cov(x, x, true) * Cov(y, y, true);
                if (d == 0) throw XlError.Div0;
                var r = Cov(x, y, true); return N(r * r / d);
            });
            Add("FORECAST.LINEAR", (c, a) =>
            {
                Need(a, 3, 3); var at = c.Num(a[0]); Pairs(c, a[1], a[2], out var y, out var x);
                if (x.Count < 2) throw XlError.Div0;
                var vx = Cov(x, x, true);
                if (vx == 0) throw XlError.Div0;
                var slope = Cov(x, y, true) / vx;
                return N(y.Average() + slope * (at - x.Average()));
            });
            Add("STEYX", (c, a) =>
            {
                Need(a, 2, 2); Pairs(c, a[0], a[1], out var y, out var x);
                if (x.Count < 3) throw XlError.Div0;
                var vx = Cov(x, x, true);
                if (vx == 0) throw XlError.Div0;
                var slope = Cov(x, y, true) / vx; var my = y.Average(); var mx = x.Average();
                double sse = 0;
                for (var i = 0; i < x.Count; i++) { var e = y[i] - (my + slope * (x[i] - mx)); sse += e * e; }
                return N(Math.Sqrt(sse / (x.Count - 2)));
            });
            Add("SKEW", (c, a) =>
            {
                var xs = c.Numbers(a); var n = xs.Count;
                if (n < 3) throw XlError.Div0;
                var m = xs.Average(); var sd = Math.Sqrt(Variance(xs, true));
                if (sd == 0) throw XlError.Div0;
                return N(n / ((n - 1.0) * (n - 2.0)) * xs.Sum(x => Math.Pow((x - m) / sd, 3)));
            });
            Add("SKEW.P", (c, a) =>
            {
                var xs = c.Numbers(a); var n = xs.Count;
                if (n < 3) throw XlError.Div0;
                var m = xs.Average(); var sd = Math.Sqrt(Variance(xs, false));
                if (sd == 0) throw XlError.Div0;
                return N(xs.Sum(x => Math.Pow((x - m) / sd, 3)) / n);
            });
            Add("KURT", (c, a) =>
            {
                var xs = c.Numbers(a); double n = xs.Count;
                if (n < 4) throw XlError.Div0;
                var m = xs.Average(); var sd = Math.Sqrt(Variance(xs, true));
                if (sd == 0) throw XlError.Div0;
                return N(n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * xs.Sum(x => Math.Pow((x - m) / sd, 4)) - 3.0 * (n - 1) * (n - 1) / ((n - 2) * (n - 3)));
            });
            Num1("GAMMALN", x => x <= 0 ? throw XlError.Num : Special.LogGamma(x));
            Num1("GAMMALN.PRECISE", x => x <= 0 ? throw XlError.Num : Special.LogGamma(x));
            Num1("GAMMA", x => x == Math.Floor(x) && x <= 0 ? throw XlError.Num : Special.Gamma(x));
            Num1("FISHER", x => Math.Abs(x) >= 1 ? throw XlError.Num : 0.5 * Math.Log((1 + x) / (1 - x)));
            Num1("FISHERINV", y => Math.Tanh(y));
        }

        static ValueEval Rank(Ctx c, ValueEval[] a, bool average)
        {
            Need(a, 2, 3);
            var x = c.Num(a[0]); var xs = c.Numbers(new[] { a[1] }); var ascending = c.Num(a, 2, 0) != 0;
            if (!xs.Contains(x)) throw XlError.NA;
            var better = xs.Count(v => ascending ? v < x : v > x);
            var ties = xs.Count(v => v == x);
            return N(average ? better + (ties + 1) / 2.0 : better + 1);
        }
    }
}
