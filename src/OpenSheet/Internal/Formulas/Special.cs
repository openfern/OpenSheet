using System;

namespace OpenSheet.Internal.Formulas
{
    /// <summary>Special functions behind the statistical worksheet functions. Accuracy is about 1e-14, which is what Excel prints.</summary>
    static class Special
    {
        const double Eps = 1e-16;
        const double Tiny = 1e-300;

        static readonly double[] Lanczos =
        {
            0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
            12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
        };

        public static double LogGamma(double x)
        {
            if (x < 0.5) return Math.Log(Math.PI / Math.Abs(Math.Sin(Math.PI * x))) - LogGamma(1 - x);
            x -= 1;
            var a = Lanczos[0];
            var t = x + 7.5;
            for (var i = 1; i < 9; i++) a += Lanczos[i] / (x + i);
            return 0.5 * Math.Log(2 * Math.PI) + (x + 0.5) * Math.Log(t) - t + Math.Log(a);
        }

        public static double Gamma(double x)
        {
            if (x == Math.Floor(x) && x > 0 && x < 171)
            {
                double r = 1;
                for (var i = 2; i < x; i++) r *= i;
                return r;
            }
            if (x < 0.5) return Math.PI / (Math.Sin(Math.PI * x) * Gamma(1 - x));
            return Math.Exp(LogGamma(x));
        }

        /// <summary>Regularized lower incomplete gamma P(a, x).</summary>
        public static double GammaP(double a, double x)
        {
            if (x <= 0) return 0;
            if (double.IsPositiveInfinity(x)) return 1;
            return x < a + 1 ? GammaSeries(a, x) : 1 - GammaFraction(a, x);
        }

        public static double GammaQ(double a, double x)
        {
            if (x <= 0) return 1;
            if (double.IsPositiveInfinity(x)) return 0;
            return x < a + 1 ? 1 - GammaSeries(a, x) : GammaFraction(a, x);
        }

        static double GammaSeries(double a, double x)
        {
            var ap = a; var sum = 1 / a; var del = sum;
            for (var n = 0; n < 1000; n++)
            {
                ap += 1; del *= x / ap; sum += del;
                if (Math.Abs(del) < Math.Abs(sum) * Eps) break;
            }
            return sum * Math.Exp(-x + a * Math.Log(x) - LogGamma(a));
        }

        static double GammaFraction(double a, double x)
        {
            var b = x + 1 - a; var c = 1 / Tiny; var d = 1 / b; var h = d;
            for (var i = 1; i < 1000; i++)
            {
                var an = -i * (i - a);
                b += 2;
                d = an * d + b; if (Math.Abs(d) < Tiny) d = Tiny;
                c = b + an / c; if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d;
                var del = d * c; h *= del;
                if (Math.Abs(del - 1) < Eps) break;
            }
            return Math.Exp(-x + a * Math.Log(x) - LogGamma(a)) * h;
        }

        public static double Erf(double x) => x >= 0 ? GammaP(0.5, x * x) : -GammaP(0.5, x * x);
        public static double Erfc(double x) => x >= 0 ? GammaQ(0.5, x * x) : 1 + GammaP(0.5, x * x);

        /// <summary>Regularized incomplete beta I_x(a, b).</summary>
        public static double BetaI(double a, double b, double x)
        {
            if (x <= 0) return 0;
            if (x >= 1) return 1;
            var bt = Math.Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b) + a * Math.Log(x) + b * Math.Log(1 - x));
            return x < (a + 1) / (a + b + 2) ? bt * BetaFraction(a, b, x) / a : 1 - bt * BetaFraction(b, a, 1 - x) / b;
        }

        static double BetaFraction(double a, double b, double x)
        {
            var qab = a + b; var qap = a + 1; var qam = a - 1;
            double c = 1, d = 1 - qab * x / qap;
            if (Math.Abs(d) < Tiny) d = Tiny;
            d = 1 / d; var h = d;
            for (var m = 1; m <= 1000; m++)
            {
                var m2 = 2 * m;
                var aa = m * (b - m) * x / ((qam + m2) * (a + m2));
                d = 1 + aa * d; if (Math.Abs(d) < Tiny) d = Tiny;
                c = 1 + aa / c; if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d; h *= d * c;
                aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
                d = 1 + aa * d; if (Math.Abs(d) < Tiny) d = Tiny;
                c = 1 + aa / c; if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d;
                var del = d * c; h *= del;
                if (Math.Abs(del - 1) < Eps) break;
            }
            return h;
        }

        public static double NormCdf(double z) => 0.5 * Erfc(-z / Math.Sqrt(2));
        public static double NormPdf(double z) => Math.Exp(-0.5 * z * z) / Math.Sqrt(2 * Math.PI);

        /// <summary>Inverse of the standard normal CDF (Wichura's AS 241, about 1e-16).</summary>
        public static double NormInv(double p)
        {
            if (p <= 0 || p >= 1) throw XlError.Num;
            var q = p - 0.5;
            double r, val;
            if (Math.Abs(q) <= 0.425)
            {
                r = 0.180625 - q * q;
                val = q * (((((((r * 2509.0809287301226727 + 33430.575583588128105) * r + 67265.770927008700853) * r + 45921.953931549871457) * r + 13731.693765509461125) * r + 1971.5909503065514427) * r + 133.14166789178437745) * r + 3.387132872796366608)
                      / (((((((r * 5226.495278852545925 + 28729.085735721942674) * r + 39307.89580009271061) * r + 21213.794301586595867) * r + 5394.1960214247511077) * r + 687.1870074920579083) * r + 42.313330701600911252) * r + 1);
                return val;
            }
            r = q < 0 ? p : 1 - p;
            r = Math.Sqrt(-Math.Log(r));
            if (r <= 5)
            {
                r -= 1.6;
                val = (((((((r * 7.7454501427834140764e-4 + 0.0227238449892691845833) * r + 0.24178072517745061177) * r + 1.27045825245236838258) * r + 3.64784832476320460504) * r + 5.7694972214606914055) * r + 4.6303378461565452959) * r + 1.42343711074968357734)
                      / (((((((r * 1.05075007164441684324e-9 + 5.475938084995344946e-4) * r + 0.0151986665636164571966) * r + 0.14810397642748007459) * r + 0.68976733498510000455) * r + 1.6763848301838038494) * r + 2.05319162663775882187) * r + 1);
            }
            else
            {
                r -= 5;
                val = (((((((r * 2.01033439929228813265e-7 + 2.71155556874348757815e-5) * r + 0.0012426609473880784386) * r + 0.026532189526576123093) * r + 0.29656057182850489123) * r + 1.7848265399172913358) * r + 5.4637849111641143699) * r + 6.6579046435011037772)
                      / (((((((r * 2.04426310338993978564e-15 + 1.4215117583164458887e-7) * r + 1.8463183175100546818e-5) * r + 7.868691311456132591e-4) * r + 0.0148753612908506148525) * r + 0.13692988092273580531) * r + 0.59983220655588793769) * r + 1);
            }
            return q < 0 ? -val : val;
        }

        /// <summary>Finds x in [lo, hi] where an increasing function reaches target (bisection; slow but safe).</summary>
        public static double Invert(Func<double, double> increasing, double target, double lo, double hi)
        {
            for (var i = 0; i < 200; i++)
            {
                var mid = 0.5 * (lo + hi);
                if (increasing(mid) < target) lo = mid; else hi = mid;
                if (Math.Abs(hi - lo) <= 1e-15 * Math.Max(1, Math.Abs(mid))) break;
            }
            return 0.5 * (lo + hi);
        }

        public static double LogChoose(double n, double k) => LogGamma(n + 1) - LogGamma(k + 1) - LogGamma(n - k + 1);

        /// <summary>Student's t cumulative distribution with <paramref name="df"/> degrees of freedom.</summary>
        public static double TCdf(double t, double df)
        {
            var x = df / (df + t * t);
            var tail = 0.5 * BetaI(df / 2, 0.5, x);
            return t > 0 ? 1 - tail : tail;
        }

        public static double ChiSqCdf(double x, double df) => GammaP(df / 2, x / 2);

        public static double FCdf(double x, double d1, double d2) => x <= 0 ? 0 : BetaI(d1 / 2, d2 / 2, d1 * x / (d1 * x + d2));
    }
}
