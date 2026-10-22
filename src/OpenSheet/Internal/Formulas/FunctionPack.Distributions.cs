using System;
using System.Linq;
using NPOI.SS.Formula.Eval;

namespace OpenSheet.Internal.Formulas
{
    static partial class FunctionPack
    {
        // Numbers 0..n-1 of the arguments plus an optional trailing TRUE/FALSE.
        static void Dist(string name, int numbers, Func<double[], bool, double> f, bool hasCumulative = true)
        {
            Add(name, (c, a) =>
            {
                Need(a, numbers + (hasCumulative ? 1 : 0), numbers + (hasCumulative ? 1 : 0) + 0);
                var x = new double[numbers];
                for (var i = 0; i < numbers; i++) x[i] = c.Num(a[i]);
                return N(f(x, hasCumulative && c.Bool(a[numbers])));
            });
        }

        static void Alias(string alias, string existing) => Map[alias] = Map[existing];

        static double LogPmfBinom(double k, double n, double p)
        {
            if (p == 0) return k == 0 ? 0 : double.NegativeInfinity;
            if (p == 1) return k == n ? 0 : double.NegativeInfinity;
            return Special.LogChoose(n, k) + k * Math.Log(p) + (n - k) * Math.Log(1 - p);
        }

        static double BinomCdf(double k, double n, double p) => k >= n ? 1 : Special.BetaI(n - k, k + 1, 1 - p);

        static void DefineDistributions()
        {
            Dist("NORM.S.DIST", 1, (x, cum) => cum ? Special.NormCdf(x[0]) : Special.NormPdf(x[0]));
            Dist("NORM.DIST", 3, (x, cum) =>
            {
                if (x[2] <= 0) throw XlError.Num;
                var z = (x[0] - x[1]) / x[2];
                return cum ? Special.NormCdf(z) : Special.NormPdf(z) / x[2];
            });
            Alias("NORMDIST", "NORM.DIST");
            Dist("NORM.INV", 3, (x, _) => x[0] <= 0 || x[0] >= 1 || x[2] <= 0 ? throw XlError.Num : x[1] + x[2] * Special.NormInv(x[0]), hasCumulative: false);
            Alias("NORMINV", "NORM.INV");
            Dist("NORM.S.INV", 1, (x, _) => Special.NormInv(x[0]), hasCumulative: false);
            Alias("NORMSINV", "NORM.S.INV");
            Num1("GAUSS", z => Special.NormCdf(z) - 0.5);
            Num1("PHI", z => Special.NormPdf(z));
            Add("CONFIDENCE.NORM", (c, a) =>
            {
                Need(a, 3, 3);
                double alpha = c.Num(a[0]), sd = c.Num(a[1]), n = Math.Truncate(c.Num(a[2]));
                if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) throw XlError.Num;
                return N(Special.NormInv(1 - alpha / 2) * sd / Math.Sqrt(n));
            });
            Alias("CONFIDENCE", "CONFIDENCE.NORM");
            Add("CONFIDENCE.T", (c, a) =>
            {
                Need(a, 3, 3);
                double alpha = c.Num(a[0]), sd = c.Num(a[1]), n = Math.Truncate(c.Num(a[2]));
                if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) throw XlError.Num;
                if (n == 1) throw XlError.Div0;
                return N(TInv(1 - alpha / 2, n - 1) * sd / Math.Sqrt(n));
            });

            Dist("BINOM.DIST", 3, (x, cum) =>
            {
                double k = Math.Truncate(x[0]), n = Math.Truncate(x[1]), p = x[2];
                if (k < 0 || k > n || p < 0 || p > 1 || n < 0) throw XlError.Num;
                return cum ? BinomCdf(k, n, p) : Math.Exp(LogPmfBinom(k, n, p));
            });
            Alias("BINOMDIST", "BINOM.DIST");
            Add("BINOM.INV", BinomInv);
            Add("CRITBINOM", BinomInv);
            Dist("POISSON.DIST", 2, (x, cum) =>
            {
                double k = Math.Truncate(x[0]), m = x[1];
                if (k < 0 || m < 0) throw XlError.Num;
                if (m == 0) return k == 0 || cum ? 1 : 0;
                return cum ? Special.GammaQ(k + 1, m) : Math.Exp(-m + k * Math.Log(m) - Special.LogGamma(k + 1));
            });
            Alias("POISSON", "POISSON.DIST");
            Dist("EXPON.DIST", 2, (x, cum) =>
            {
                if (x[0] < 0 || x[1] <= 0) throw XlError.Num;
                return cum ? 1 - Math.Exp(-x[1] * x[0]) : x[1] * Math.Exp(-x[1] * x[0]);
            });
            Alias("EXPONDIST", "EXPON.DIST");
            Dist("NEGBINOM.DIST", 3, (x, cum) =>
            {
                double f = Math.Truncate(x[0]), s = Math.Truncate(x[1]), p = x[2];
                if (f < 0 || s < 1 || p <= 0 || p >= 1) throw XlError.Num;   // 0 and 1 are refused too
                return cum ? Special.BetaI(s, f + 1, p) : Math.Exp(Special.LogChoose(f + s - 1, f) + s * Math.Log(p) + f * Math.Log(1 - p));
            });
            // the pre-2010 form has no "cumulative" argument: it is always the probability of exactly that many
            Add("NEGBINOMDIST", (c, a) =>
            {
                Need(a, 3, 3);
                double f = Math.Truncate(c.Num(a[0])), s = Math.Truncate(c.Num(a[1])), p = c.Num(a[2]);
                if (f < 0 || s < 1 || p <= 0 || p >= 1) throw XlError.Num;   // 0 and 1 are refused too
                return N(Math.Exp(Special.LogChoose(f + s - 1, f) + s * Math.Log(p) + f * Math.Log(1 - p)));
            });
            Add("HYPGEOM.DIST", (c, a) => { Need(a, 5, 5); return Hypergeometric(c, a, c.Bool(a[4])); });
            Add("HYPGEOMDIST", (c, a) => { Need(a, 4, 4); return Hypergeometric(c, a, false); });

            // Student's t
            Dist("T.DIST", 2, (x, cum) =>
            {
                var df = Df(Math.Truncate(x[1]));
                if (df < 1) throw XlError.Num;
                if (cum) return Special.TCdf(x[0], df);
                return Math.Exp(Special.LogGamma((df + 1) / 2) - Special.LogGamma(df / 2) - 0.5 * Math.Log(df * Math.PI) - (df + 1) / 2 * Math.Log(1 + x[0] * x[0] / df));
            });
            Add("T.DIST.2T", (c, a) => { Need(a, 2, 2); var x = c.Num(a[0]); var df = Df(Math.Truncate(c.Num(a[1]))); if (x < 0 || df < 1) throw XlError.Num; return N(2 * (1 - Special.TCdf(x, df))); });
            Add("T.DIST.RT", (c, a) => { Need(a, 2, 2); var x = c.Num(a[0]); var df = Df(Math.Truncate(c.Num(a[1]))); if (df < 1) throw XlError.Num; return N(1 - Special.TCdf(x, df)); });
            Add("T.INV", (c, a) => { Need(a, 2, 2); var p = c.Num(a[0]); var df = Math.Truncate(c.Num(a[1])); if (p <= 0 || p >= 1 || df < 1) throw XlError.Num; return N(TInv(p, df)); });
            Add("T.INV.2T", (c, a) => { Need(a, 2, 2); var p = c.Num(a[0]); var df = Math.Truncate(c.Num(a[1])); if (p <= 0 || p > 1 || df < 1) throw XlError.Num; return N(TInv(1 - p / 2, df)); });

            // chi-square
            Dist("CHISQ.DIST", 2, (x, cum) =>
            {
                var df = Df(Math.Truncate(x[1]));
                if (x[0] < 0 || df < 1) throw XlError.Num;
                return cum ? Special.ChiSqCdf(x[0], df) : Math.Exp((df / 2 - 1) * Math.Log(x[0]) - x[0] / 2 - df / 2 * Math.Log(2) - Special.LogGamma(df / 2));
            });
            Add("CHISQ.DIST.RT", (c, a) => { Need(a, 2, 2); var x = c.Num(a[0]); var df = Df(Math.Truncate(c.Num(a[1]))); if (x < 0 || df < 1) throw XlError.Num; return N(1 - Special.ChiSqCdf(x, df)); });
            Alias("CHIDIST", "CHISQ.DIST.RT");
            Add("CHISQ.INV", (c, a) => { Need(a, 2, 2); var p = c.Num(a[0]); var df = Math.Truncate(c.Num(a[1])); if (p < 0 || p >= 1 || df < 1) throw XlError.Num; return N(ChiInv(p, df)); });
            Add("CHISQ.INV.RT", (c, a) => { Need(a, 2, 2); var p = c.Num(a[0]); var df = Math.Truncate(c.Num(a[1])); if (p <= 0 || p > 1 || df < 1) throw XlError.Num; return N(ChiInv(1 - p, df)); });
            Alias("CHIINV", "CHISQ.INV.RT");

            // F
            Dist("F.DIST", 3, (x, cum) =>
            {
                double d1 = Df(Math.Truncate(x[1])), d2 = Df(Math.Truncate(x[2]));
                if (x[0] < 0 || d1 < 1 || d2 < 1) throw XlError.Num;
                if (cum) return Special.FCdf(x[0], d1, d2);
                if (x[0] == 0) return d1 > 2 ? 0 : d1 == 2 ? 1 : double.PositiveInfinity;
                return Math.Exp(0.5 * (d1 * Math.Log(d1 * x[0]) + d2 * Math.Log(d2) - (d1 + d2) * Math.Log(d1 * x[0] + d2)) - Math.Log(x[0]) - (Special.LogGamma(d1 / 2) + Special.LogGamma(d2 / 2) - Special.LogGamma((d1 + d2) / 2)));
            });
            Add("F.DIST.RT", (c, a) => { Need(a, 3, 3); double x = c.Num(a[0]), d1 = Df(Math.Truncate(c.Num(a[1]))), d2 = Df(Math.Truncate(c.Num(a[2]))); if (x < 0 || d1 < 1 || d2 < 1) throw XlError.Num; return N(1 - Special.FCdf(x, d1, d2)); });
            Alias("FDIST", "F.DIST.RT");
            Add("F.INV", (c, a) => { Need(a, 3, 3); double p = c.Num(a[0]), d1 = Math.Truncate(c.Num(a[1])), d2 = Math.Truncate(c.Num(a[2])); if (p < 0 || p >= 1 || d1 < 1 || d2 < 1) throw XlError.Num; return N(p == 0 ? 0 : FInv(p, d1, d2)); });
            Add("F.INV.RT", (c, a) => { Need(a, 3, 3); double p = c.Num(a[0]), d1 = Math.Truncate(c.Num(a[1])), d2 = Math.Truncate(c.Num(a[2])); if (p <= 0 || p > 1 || d1 < 1 || d2 < 1) throw XlError.Num; return N(p == 1 ? 0 : FInv(1 - p, d1, d2)); });
            Alias("FINV", "F.INV.RT");

            // continuous families
            Dist("LOGNORM.DIST", 3, (x, cum) =>
            {
                if (x[0] <= 0 || x[2] <= 0) throw XlError.Num;
                var z = (Math.Log(x[0]) - x[1]) / x[2];
                return cum ? Special.NormCdf(z) : Special.NormPdf(z) / (x[0] * x[2]);
            });
            Add("LOGNORMDIST", (c, a) =>   // the pre-2010 form has no "cumulative" argument: it is always the cumulative probability
            {
                Need(a, 3, 3);
                double x = c.Num(a[0]), m = c.Num(a[1]), sd = c.Num(a[2]);
                if (x <= 0 || sd <= 0) throw XlError.Num;
                return N(Special.NormCdf((Math.Log(x) - m) / sd));
            });
            Add("LOGNORM.INV", (c, a) => { Need(a, 3, 3); double p = c.Num(a[0]), m = c.Num(a[1]), sd = c.Num(a[2]); if (p <= 0 || p >= 1 || sd <= 0) throw XlError.Num; return N(Math.Exp(m + sd * Special.NormInv(p))); });
            Alias("LOGINV", "LOGNORM.INV");
            Dist("WEIBULL.DIST", 3, (x, cum) =>
            {
                if (x[0] < 0 || x[1] <= 0 || x[2] <= 0) throw XlError.Num;
                var t = Math.Pow(x[0] / x[2], x[1]);
                return cum ? 1 - Math.Exp(-t) : x[1] / x[2] * Math.Pow(x[0] / x[2], x[1] - 1) * Math.Exp(-t);
            });
            Alias("WEIBULL", "WEIBULL.DIST");
            Dist("GAMMA.DIST", 3, (x, cum) =>
            {
                if (x[0] < 0 || x[1] <= 0 || x[2] <= 0) throw XlError.Num;
                return cum ? Special.GammaP(x[1], x[0] / x[2]) : Math.Exp((x[1] - 1) * Math.Log(x[0]) - x[0] / x[2] - x[1] * Math.Log(x[2]) - Special.LogGamma(x[1]));
            });
            Alias("GAMMADIST", "GAMMA.DIST");
            Add("GAMMA.INV", (c, a) =>
            {
                Need(a, 3, 3); double p = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]);
                if (p < 0 || p >= 1 || al <= 0 || be <= 0) throw XlError.Num;
                if (p == 0) return N(0);
                return N(be * Special.Invert(v => Special.GammaP(al, v), p, 0, Math.Max(10, al * 20)));
            });
            Alias("GAMMAINV", "GAMMA.INV");
            Add("BETA.DIST", (c, a) =>
            {
                Need(a, 4, 6);
                double x = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]); var cum = c.Bool(a[3]);
                double lo = c.Num(a, 4, 0), hi = c.Num(a, 5, 1);
                if (al <= 0 || be <= 0 || x < lo || x > hi || lo >= hi) throw XlError.Num;
                var t = (x - lo) / (hi - lo);
                if (cum) return N(Special.BetaI(al, be, t));
                return N(Math.Exp((al - 1) * Math.Log(t) + (be - 1) * Math.Log(1 - t) + Special.LogGamma(al + be) - Special.LogGamma(al) - Special.LogGamma(be)) / (hi - lo));
            });
            Add("BETA.INV", (c, a) =>
            {
                Need(a, 3, 5);
                double p = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]), lo = c.Num(a, 3, 0), hi = c.Num(a, 4, 1);
                if (p <= 0 || p >= 1 || al <= 0 || be <= 0 || lo >= hi) throw XlError.Num;
                return N(lo + (hi - lo) * Special.Invert(v => Special.BetaI(al, be, v), p, 0, 1));
            });
            Alias("BETAINV", "BETA.INV");
            Add("BETADIST", (c, a) =>
            {
                Need(a, 3, 5);
                double x = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]), lo = c.Num(a, 3, 0), hi = c.Num(a, 4, 1);
                if (al <= 0 || be <= 0 || x < lo || x > hi || lo >= hi) throw XlError.Num;
                return N(Special.BetaI(al, be, (x - lo) / (hi - lo)));
            });
        }

        // sample successes k out of n drawn, kk successes in a population of nn
        static ValueEval Hypergeometric(Ctx c, ValueEval[] a, bool cum)
        {
            double k = Math.Truncate(c.Num(a[0])), n = Math.Truncate(c.Num(a[1])), kk = Math.Truncate(c.Num(a[2])), nn = Math.Truncate(c.Num(a[3]));
            if (n < 0 || n > nn || kk < 0 || kk > nn || nn <= 0 || k < 0 || k > Math.Min(n, kk)) throw XlError.Num;
            if (k < Math.Max(0, n - nn + kk)) return N(0);   // an impossible sample has probability 0, it is not an error
            double Pmf(double j) => Math.Exp(Special.LogChoose(kk, j) + Special.LogChoose(nn - kk, n - j) - Special.LogChoose(nn, n));
            if (!cum) return N(Pmf(k));
            double sum = 0;
            for (var j = Math.Max(0, n - nn + kk); j <= k; j++) sum += Pmf(j);
            return N(sum);
        }

        // Excel refuses 10^10 degrees of freedom or more
        static double Df(double df) => df >= 1e10 ? throw XlError.Num : df;

        static double TInv(double p, double df)
        {
            if (p == 0.5) return 0;
            var upper = p > 0.5;
            var q = upper ? 1 - p : p;
            // find t > 0 with right tail q, then mirror
            var t = Special.Invert(v => 1 - 0.5 * Special.BetaI(df / 2, 0.5, df / (df + v * v)), 1 - q, 0, 1e8);
            return upper ? t : -t;
        }

        static double ChiInv(double p, double df) => p == 0 ? 0 : Special.Invert(v => Special.ChiSqCdf(v, df), p, 0, Math.Max(100, df * 50));

        static double FInv(double p, double d1, double d2) => Special.Invert(v => Special.FCdf(v, d1, d2), p, 0, 1e9);

        static ValueEval BinomInv(Ctx c, ValueEval[] a)
        {
            Need(a, 3, 3);
            double n = Math.Truncate(c.Num(a[0])), p = c.Num(a[1]), alpha = c.Num(a[2]);
            if (n < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1) throw XlError.Num;
            for (double k = 0; k <= n; k++) if (BinomCdf(k, n, p) >= alpha) return N(k);
            return N(n);
        }
    }
}
