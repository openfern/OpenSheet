using System;
using System.Globalization;
using System.Linq;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal.Formulas
{
    static partial class FunctionPack
    {
        static DateTime FromSerial(double serial)
        {
            if (serial < 0 || serial >= 2958466) throw XlError.Num;
            return XL.DateUtil.GetJavaDate(Math.Floor(serial));
        }

        static string SheetNameOrNull(IEvaluationWorkbook book, int index)
        {
            try { return book.GetSheetName(index); }
            catch (Exception) { return null; }
        }

        // The cell a reference points at (the first one of a range), or null when there is no such cell; a value that is no reference is #VALUE!.
        static IEvaluationCell CellOf(Ctx c, ValueEval v)
        {
            if (c.Workbook == null) throw XlError.NA;
            int row, column;
            if (v is RefEval r) { row = r.Row; column = r.Column; }
            else if (v is AreaEval area) { row = area.FirstRow; column = area.FirstColumn; }
            else throw XlError.Value;
            var sheet = v is NPOI.SS.Formula.ISheetRange span ? span.FirstSheetIndex : c.SheetIndex;
            return c.Workbook.GetSheet(sheet < 0 ? 0 : sheet)?.GetCell(row, column);
        }

        // 0 = Sunday, as Excel counts them: serial 1 is a Sunday even though 1 January 1900 was a Monday
        static int Dow(double serial)
        {
            if (serial < 0 || serial >= 2958466) throw XlError.Num;
            return (int)((long)Math.Floor(serial) + 6) % 7;
        }

        static void DefineInfo()
        {
            Add("XOR", (c, a) =>
            {
                var count = 0; var any = false;
                foreach (var arg in a)
                {
                    if (arg is AreaEval)
                    {
                        foreach (var x in c.Area(arg).Flat())
                        {
                            if (x is ErrorEval e) throw new XlError(e);
                            if (x is BoolEval b) { any = true; if (b.BooleanValue) count++; }
                            else if (x is NumberEval n) { any = true; if (n.NumberValue != 0) count++; }
                        }
                    }
                    else if (arg is StringEval || (arg is RefEval && c.Scalar(arg) is BlankEval)) continue;   // text typed into the formula and empty cells are ignored, as inside a range
                    else { any = true; if (c.Bool(arg)) count++; }
                }
                if (!any) throw XlError.Value;
                return B(count % 2 == 1);
            });
            Add("N", (c, a) =>
            {
                Need(a, 1, 1);
                switch (c.Scalar(a[0]))
                {
                    case NumberEval n: return N(n.NumberValue);
                    case BoolEval b: return N(b.BooleanValue ? 1 : 0);
                    case ErrorEval e: return e;
                    default: return N(0);
                }
            });
            Add("SHEET", (c, a) =>
            {
                Need(a, 0, 1);
                if (a.Length == 1 && a[0] is NPOI.SS.Formula.ISheetRange r) return N(r.FirstSheetIndex + 1);
                if (a.Length == 1 && a[0] is StringEval name && c.Workbook != null)   // a sheet's name
                {
                    var found = c.Workbook.GetSheetIndex(name.StringValue);
                    if (found < 0) throw XlError.NA;
                    return N(found + 1);
                }
                if (a.Length == 1 && (a[0] is NumberEval || a[0] is BoolEval || a[0] is ErrorEval || a[0] is BlankEval)) throw XlError.NA;   // neither a reference nor a sheet name
                if (c.SheetIndex < 0) throw XlError.NA;
                return N(c.SheetIndex + 1);
            });
            Add("SHEETS", (c, a) =>
            {
                Need(a, 0, 1);
                if (a.Length == 1)   // a reference (even across sheets) counts the sheets it spans; anything else is not one
                {
                    if (a[0] is NPOI.SS.Formula.ISheetRange span) return N(span.LastSheetIndex - span.FirstSheetIndex + 1);
                    throw XlError.NA;
                }
                if (c.Workbook == null) throw XlError.NA;
                var count = 0;
                while (!string.IsNullOrEmpty(SheetNameOrNull(c.Workbook, count))) count++;
                return N(count);
            });
            Add("ISFORMULA", (c, a) =>
            {
                Need(a, 1, 1);
                return B(CellOf(c, a[0])?.CellType == NPOI.SS.UserModel.CellType.Formula);
            });
            Add("FORMULATEXT", (c, a) =>
            {
                Need(a, 1, 1);
                var cell = CellOf(c, a[0]);
                if (cell == null || cell.CellType != NPOI.SS.UserModel.CellType.Formula) throw XlError.NA;
                var tokens = c.Workbook.GetFormulaTokens(cell);
                if (!(c.Workbook is NPOI.SS.Formula.IFormulaRenderingWorkbook rendering)) throw XlError.NA;
                return S("=" + FormulaDialect.ToDisplay(NPOI.SS.Formula.FormulaRenderer.ToFormulaString(rendering, tokens)));
            });
            Add("TYPE", (c, a) =>
            {
                Need(a, 1, 1);
                if (a[0] is AreaEval) return N(64);
                if (a[0] is ErrorEval) return N(16);   // an error is the thing being asked about, not a failure
                if (a[0] is RefEval held && held.GetInnerValueEval(held.FirstSheetIndex) is ErrorEval) return N(16);
                switch (c.Scalar(a[0]))
                {
                    case NumberEval _: return N(1);
                    case StringEval _: return N(2);
                    case BoolEval _: return N(4);
                    case ErrorEval _: return N(16);
                    default: return N(1);
                }
            });
        }

        static void DefineDates()
        {
            Add("DAYS", (c, a) => { Need(a, 2, 2); double end = Math.Truncate(c.Num(a[0])), start = Math.Truncate(c.Num(a[1])); Dow(end); Dow(start); return N(end - start); });
            Add("ISOWEEKNUM", (c, a) =>
            {
                Need(a, 1, 1);
                var serial = Math.Floor(c.Num(a[0]));
                if (serial < 61)   // Excel's calendar before March 1900: serial 1 is a Sunday and February has a 29th
                {
                    FromSerial(serial);
                    var isoDay = (Dow(serial) + 6) % 7 + 1;   // Monday = 1 .. Sunday = 7
                    var week = ((int)Math.Max(serial, 0) - isoDay + 10) / 7;
                    return N(week < 1 ? 52 : week);
                }
                return N(IsoWeek(FromSerial(serial)));
            });
            Add("DATEDIF", (c, a) =>
            {
                Need(a, 3, 3);
                double s = Math.Floor(c.Num(a[0])), e = Math.Floor(c.Num(a[1]));
                var unit = c.Str(a[2]).ToUpperInvariant();
                if (s > e) throw XlError.Num;
                DateTime ds = FromSerial(s), de = FromSerial(e);
                var months = (de.Year - ds.Year) * 12 + de.Month - ds.Month - (de.Day < ds.Day ? 1 : 0);
                switch (unit)
                {
                    case "Y": return N(months / 12);
                    case "M": return N(months);
                    case "D": return N(e - s);
                    case "MD":
                        if (de.Day >= ds.Day) return N(de.Day - ds.Day);
                        var prev = de.AddMonths(-1);
                        return N(DateTime.DaysInMonth(prev.Year, prev.Month) - ds.Day + de.Day);
                    case "YM": return N(((months % 12) + 12) % 12);
                    case "YD":
                        var anniversary = SafeDate(de.Year, ds.Month, ds.Day);
                        if (anniversary > de) anniversary = SafeDate(de.Year - 1, ds.Month, ds.Day);
                        return N((de - anniversary).TotalDays);
                    default: throw XlError.Num;
                }
            });
            Add("TIMEVALUE", (c, a) =>
            {
                Need(a, 1, 1);
                if (!(c.Scalar(a[0]) is StringEval)) throw XlError.Value;   // a number is not a time written as text
                var t = c.Str(a[0]).Trim();
                if (t.IndexOf(':') < 0 && t.IndexOf("AM", StringComparison.OrdinalIgnoreCase) < 0 && t.IndexOf("PM", StringComparison.OrdinalIgnoreCase) < 0) throw XlError.Value;
                if (DateTime.TryParse(t, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowWhiteSpaces, out var d)) return N(d.TimeOfDay.TotalDays);
                if (TimeSpan.TryParse(t, CultureInfo.InvariantCulture, out var ts)) return N(ts.TotalDays - Math.Floor(ts.TotalDays));
                throw XlError.Value;
            });
            Add("NETWORKDAYS.INTL", NetworkdaysIntl);
        }

        static ValueEval NetworkdaysIntl(Ctx c, ValueEval[] a)
        {
            Need(a, 2, 4);
            double s = Math.Floor(c.Num(a[0])), e = Math.Floor(c.Num(a[1]));
            var weekend = Weekend(c, a, 2);
            var holidays = Holidays(c, a, 3);
            var sign = s <= e ? 1 : -1;
            var lo = Math.Min(s, e); var hi = Math.Max(s, e);
            var n = 0;
            for (var d = lo; d <= hi; d++)
                if (!weekend[Dow(d)] && !holidays.Contains(d)) n++;
            return N(sign * n);
        }

        // ISO 8601 week: the week containing the year's first Thursday is week 1.
        static int IsoWeek(DateTime d)
        {
            var day = d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
            return (d.AddDays(4 - day).DayOfYear - 1) / 7 + 1;
        }

        static DateTime SafeDate(int year, int month, int day) => new DateTime(year, month, Math.Min(day, DateTime.DaysInMonth(year, month)));

        // true for each day of the week (Sunday = 0) that is a weekend day
        static bool[] Weekend(Ctx c, ValueEval[] a, int index)
        {
            var w = new bool[7];
            if (index >= a.Length || Ctx.IsMissing(a[index])) { w[0] = w[6] = true; return w; }
            var v = c.Scalar(a[index]);
            if (v is StringEval s)
            {
                if (s.StringValue.Length != 7 || s.StringValue.Any(ch => ch != '0' && ch != '1') || s.StringValue == "1111111") throw XlError.Value;
                for (var i = 0; i < 7; i++) w[(i + 1) % 7] = s.StringValue[i] == '1'; // the string starts on Monday
                return w;
            }
            var code = (int)Math.Truncate(c.Num(v));
            if (code >= 1 && code <= 7) { w[(code + 5) % 7] = true; w[(code + 6) % 7] = true; }       // 1 = Sat+Sun, 2 = Sun+Mon, ...
            else if (code >= 11 && code <= 17) w[code - 11] = true;                                    // 11 = Sunday only, 12 = Monday only, ...
            else throw XlError.Num;
            return w;
        }

        static System.Collections.Generic.HashSet<double> Holidays(Ctx c, ValueEval[] a, int index)
        {
            var set = new System.Collections.Generic.HashSet<double>();
            if (index < a.Length && !Ctx.IsMissing(a[index]))
                foreach (var x in c.Area(a[index]).Flat())
                {
                    if (x is NumberEval n) { if (n.NumberValue < 0) throw XlError.Num; set.Add(Math.Floor(n.NumberValue)); }
                    else if (x is ErrorEval e) throw new XlError(e);
                    else if (x is BoolEval) throw XlError.Value;
                    else if (x is StringEval s)   // a holiday written as text must read as a date number
                    {
                        var parsed = Ctx.ParseNumber(s.StringValue) ?? throw XlError.Value;
                        if (parsed < 0) throw XlError.Num;
                        set.Add(Math.Floor(parsed));
                    }
                }
            return set;
        }
    }
}
