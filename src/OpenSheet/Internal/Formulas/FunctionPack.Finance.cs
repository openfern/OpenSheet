using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Formula.Eval;

namespace OpenSheet.Internal.Formulas
{
    static partial class FunctionPack
    {
        // Loan schedule with Excel's sign convention: money borrowed (pv) is positive, payments come out negative.
        static double LoanPayment(double rate, double nper, double pv, double fv, int type)
        {
            if (rate == 0) return -(pv + fv) / nper;
            var f = Math.Pow(1 + rate, nper);
            return -(pv * f + fv) * rate / ((1 + rate * type) * (f - 1));
        }

        // Interest paid with payment number `per` (negative, like the payment). With payments at the start of the period the
        // first payment carries no interest, and payment k pays what accrued over period k-1.
        static double LoanInterest(double rate, double per, double nper, double pv, double fv, int type)
        {
            var pmt = LoanPayment(rate, nper, pv, fv, type);
            if (type == 1 && per == 1) return 0;
            double balance;
            if (type == 0)
            {
                var j = per - 1; // payments made so far
                balance = rate == 0 ? pv + pmt * j : pv * Math.Pow(1 + rate, j) + pmt * (Math.Pow(1 + rate, j) - 1) / rate;
            }
            else
            {
                var j = per - 1; // balance right after payment j, before its interest accrues
                balance = rate == 0 ? pv + pmt * j : pv * Math.Pow(1 + rate, j - 1) + pmt * (Math.Pow(1 + rate, j) - 1) / rate;
            }
            return -balance * rate;
        }

        static IEnumerable<double> Series(Ctx c, ValueEval v)
        {
            if (v is StringEval) { yield return c.Num(v); yield break; }   // a single value typed into the formula may be text that reads as a number
            foreach (var x in c.Area(v).Flat())
            {
                if (x is NumberEval n) yield return n.NumberValue;
                else if (x is BlankEval) yield return 0;
                else if (x is ErrorEval e) throw new XlError(e);
                else throw XlError.Value;
            }
        }

        static void DefineFinance()
        {
            Num3("SLN", (cost, salvage, life) => life == 0 ? throw XlError.Div0 : (cost - salvage) / life);
            Add("SYD", (c, a) =>
            {
                Need(a, 4, 4);
                double cost = c.Num(a[0]), salvage = c.Num(a[1]), life = c.Num(a[2]), per = c.Num(a[3]);
                if (life <= 0 || per <= 0 || per > life) throw XlError.Num;
                return N((cost - salvage) * (life - per + 1) * 2 / (life * (life + 1)));
            });
            Add("DB", (c, a) =>
            {
                Need(a, 4, 5);
                double cost = c.Num(a[0]), salvage = c.Num(a[1]), life = c.Num(a[2]), period = c.Num(a[3]); var month = Math.Truncate(c.Num(a, 4, 12));
                if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12 || period > life + (month < 12 ? 1 : 0)) throw XlError.Num;
                if (cost == 0) return N(0);
                var rate = Math.Round(1 - Math.Pow(salvage / cost, 1 / life), 3);
                double total = 0, dep = 0;
                for (var p = 1; p <= Math.Truncate(period); p++)
                {
                    dep = p == 1 ? cost * rate * month / 12 : p == life + 1 ? (cost - total) * rate * (12 - month) / 12 : (cost - total) * rate;
                    total += dep;
                }
                return N(dep);
            });
            Add("DDB", (c, a) =>
            {
                Need(a, 4, 5);
                double cost = c.Num(a[0]), salvage = c.Num(a[1]), life = c.Num(a[2]), period = c.Num(a[3]), factor = c.Num(a, 4, 2);
                if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0) throw XlError.Num;
                double total = 0, dep = 0;
                for (var p = 1; p <= Math.Ceiling(period); p++)
                {
                    dep = Math.Min((cost - total) * (factor / life), Math.Max(0, cost - salvage - total));
                    total += dep;
                }
                return N(dep);
            });
            Num2("EFFECT", (nominal, n) => { n = Math.Truncate(n); return nominal <= 0 || n < 1 ? throw XlError.Num : Math.Pow(1 + nominal / n, n) - 1; });
            Num2("NOMINAL", (effect, n) => { n = Math.Truncate(n); return effect <= 0 || n < 1 ? throw XlError.Num : n * (Math.Pow(1 + effect, 1 / n) - 1); });
            Num3("PDURATION", (rate, pv, fv) => rate <= 0 || pv <= 0 || fv <= 0 ? throw XlError.Num : (Math.Log(fv) - Math.Log(pv)) / Math.Log(1 + rate));
            Num3("RRI", (n, pv, fv) => n <= 0 || pv == 0 ? throw XlError.Num : Math.Pow(fv / pv, 1 / n) - 1);
            Add("ISPMT", (c, a) =>
            {
                Need(a, 4, 4);
                double rate = c.Num(a[0]), per = c.Num(a[1]), nper = c.Num(a[2]), pv = c.Num(a[3]);
                if (nper == 0) throw XlError.Div0;
                return N(pv * rate * (per / nper - 1));
            });
            Add("FVSCHEDULE", (c, a) =>
            {
                Need(a, 2, 2);
                return N(Series(c, a[1]).Aggregate(c.Num(a[0]), (p, r) => p * (1 + r)));
            });
            // NPOI's IPMT and PPMT refuse the "payments at the start of the period" argument
            Correct("IPMT", (c, a) => Payment(c, a, interest: true));
            Correct("PPMT", (c, a) => Payment(c, a, interest: false));
            Add("CUMIPMT", (c, a) => Cumulative(c, a, interest: true));
            Add("CUMPRINC", (c, a) => Cumulative(c, a, interest: false));
            Add("XNPV", (c, a) =>
            {
                Need(a, 3, 3);
                var rate = c.Num(a[0]); var v = Series(c, a[1]).ToList(); var d = Series(c, a[2]).ToList();
                if (v.Count != d.Count || v.Count == 0) throw XlError.Num;
                if (rate <= 0) throw XlError.Num;   // Excel refuses a rate of 0 or less, not only -1 or less
                if (d.Any(x => x < d[0])) throw XlError.Num;
                return N(Enumerable.Range(0, v.Count).Sum(i => v[i] / Math.Pow(1 + rate, (Math.Truncate(d[i]) - Math.Truncate(d[0])) / 365)));
            });
            Add("XIRR", (c, a) =>
            {
                Need(a, 2, 3);
                var v = Series(c, a[0]).ToList(); var d = Series(c, a[1]).ToList(); var guess = c.Num(a, 2, 0.1);
                if (v.Count != d.Count || v.Count < 2 || !v.Any(x => x > 0) || !v.Any(x => x < 0) || d.Any(x => x < d[0])) throw XlError.Num;
                double F(double r) => Enumerable.Range(0, v.Count).Sum(i => v[i] / Math.Pow(1 + r, (Math.Truncate(d[i]) - Math.Truncate(d[0])) / 365));
                double r0 = guess;
                for (var i = 0; i < 100; i++)
                {
                    var f = F(r0); var h = 1e-7; var df = (F(r0 + h) - f) / h;
                    if (df == 0 || double.IsNaN(df)) break;
                    var r1 = r0 - f / df;
                    if (r1 <= -1) r1 = (r0 - 1) / 2;
                    if (Math.Abs(r1 - r0) < 1e-12) return N(r1);
                    r0 = r1;
                }
                throw XlError.Num;
            });
        }

        // IPMT / PPMT (rate, per, nper, pv, [fv], [type])
        static ValueEval Payment(Ctx c, ValueEval[] a, bool interest)
        {
            Need(a, 4, 6);
            double rate = c.Num(a[0]), per = c.Num(a[1]), nper = c.Num(a[2]), pv = c.Num(a[3]), fv = c.Num(a, 4, 0);
            var type = c.Num(a, 5, 0) != 0 ? 1 : 0;
            if (per < 1 || per > nper || nper <= 0) throw XlError.Num;
            var ipmt = LoanInterest(rate, Math.Truncate(per), nper, pv, fv, type);
            return N(interest ? ipmt : LoanPayment(rate, nper, pv, fv, type) - ipmt);
        }

        static ValueEval Cumulative(Ctx c, ValueEval[] a, bool interest)
        {
            Need(a, 6, 6);
            double rate = c.Num(a[0]), nper = c.Num(a[1]), pv = c.Num(a[2]), start = Math.Truncate(c.Num(a[3])), end = Math.Truncate(c.Num(a[4])); var type = (int)c.Num(a[5]);
            if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper || (type != 0 && type != 1)) throw XlError.Num;
            var pmt = LoanPayment(rate, nper, pv, 0, type);
            double sum = 0;
            for (var per = start; per <= end; per++)
            {
                var i = LoanInterest(rate, per, nper, pv, 0, type);
                sum += interest ? i : pmt - i;
            }
            return N(sum);
        }
    }
}
