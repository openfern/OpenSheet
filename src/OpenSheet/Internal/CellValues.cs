using System;
using System.Globalization;
using OpenSheet.Internal.Formulas;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    internal enum Kind { Empty, Number, Date, Text, Bool, Error }

    /// <summary>Reads and writes raw cell values. One place decides how .NET values map to cell types.</summary>
    internal static class CellValues
    {
        public const short DefaultDateFormat = 22; // Excel built-in "m/d/yy h:mm"

        public static bool IsDateFormatted(WorkBook book, XL.ICell c) => book.NumberStyleOf(c).IsDate;

        /// <summary>A date-formatted number is a DateTime only when it lies in the range Excel can draw as a date (0 to the year 9999).</summary>
        static bool IsDateValue(WorkBook book, XL.ICell c, double serial) => serial >= 0 && serial < 2958466 && IsDateFormatted(book, c);

        /// <summary>
        /// The index of the cell's style in the workbook's style table, read straight from the XLSX cell element so no style object is built;
        /// -1 when it cannot be known that way (XLS, or a NPOI that no longer has the accessor).
        /// </summary>
        public static int StyleIndex(XL.ICell c)
        {
            if (CtOf == null || !(c is NPOI.XSSF.UserModel.XSSFCell x)) return -1;
            var ct = CtOf(x);
            return ct == null ? -1 : (int)ct.s;
        }

        /// <summary>Returns the typed value: double, DateTime, string, bool, error text (string), or null for blank.</summary>
        public static object Read(WorkBook book, XL.ICell c, out Kind kind)
        {
            kind = Kind.Empty;
            if (c == null) return null;
            var t = c.CellType;
            if (t == XL.CellType.Formula)
            {
                try
                {
                    XL.CellValue cv;
                    using (EnterInvariant()) cv = book.Evaluator.Evaluate(c);
                    if (cv != null) return FromEvaluated(book, c, cv, out kind);
                }
                catch (Exception e) when (IsNotImplemented(e))
                {
                    // A function we cannot compute. A result stored in the file (written by Excel) is still right; for a formula
                    // that was just typed the "cached" value is a meaningless 0, so say so the way Excel does: #NAME?.
                    if (!HasStoredResult(c)) { kind = Kind.Error; return "#NAME?"; }
                }
                catch (Exception) { /* bad reference and the like: fall back to the cached result */ }
                t = c.CachedFormulaResultType;
            }
            else if (t == XL.CellType.Blank) return null;
            return FromCell(book, c, t, out kind);
        }

        /// <summary>The cell's value when it is a plain number (not a date), without boxing it; false for everything else.</summary>
        public static bool TryNumber(WorkBook book, XL.ICell c, out double value)
        {
            if (c != null && c.CellType == XL.CellType.Numeric)
            {
                value = c.NumericCellValue;
                return !(IsDateValue(book, c, value) && Serial.TryToDateTime(value, out _));
            }
            var v = Read(book, c, out var kind); // formulas and the rest take the general road
            if (kind == Kind.Number) { value = (double)v; return true; }
            value = 0;
            return false;
        }

        public static bool IsNotImplemented(Exception e)
        {
            for (; e != null; e = e.InnerException)
                if (e is NotImplementedException || e.GetType().Name == "NotImplementedException") return true;
            return false;
        }

        /// <summary>True when the file holds a computed result for this formula cell. Only XLSX can tell; XLS always has one.</summary>
        static bool HasStoredResult(XL.ICell c)
        {
            if (!(c is NPOI.XSSF.UserModel.XSSFCell)) return true;
            // NPOI keeps the XML element behind a non-public accessor (pinned version 2.7.6, covered by a test).
            var ct = CtOf?.Invoke((NPOI.XSSF.UserModel.XSSFCell)c);
            return ct == null || ct.IsSetV();
        }

        /// <summary>XSSFCell.GetCTCell as a delegate (a reflection call per cell would cost more than the work it saves); null if NPOI changes it.</summary>
        static readonly Func<NPOI.XSSF.UserModel.XSSFCell, NPOI.OpenXmlFormats.Spreadsheet.CT_Cell> CtOf = CtAccessor();

        static Func<NPOI.XSSF.UserModel.XSSFCell, NPOI.OpenXmlFormats.Spreadsheet.CT_Cell> CtAccessor()
        {
            var method = typeof(NPOI.XSSF.UserModel.XSSFCell).GetMethod("GetCTCell", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
            try { return method == null ? null : (Func<NPOI.XSSF.UserModel.XSSFCell, NPOI.OpenXmlFormats.Spreadsheet.CT_Cell>)Delegate.CreateDelegate(typeof(Func<NPOI.XSSF.UserModel.XSSFCell, NPOI.OpenXmlFormats.Spreadsheet.CT_Cell>), method); }
            catch (ArgumentException) { return null; }
        }

        /// <summary>Evaluates one formula cell and reports true when it was skipped because a function is not supported.</summary>
        public static bool IsUnsupported(XL.IFormulaEvaluator evaluator, XL.ICell cell)
        {
            try { using (EnterInvariant()) evaluator.EvaluateFormulaCell(cell); return false; }
            catch (Exception e) when (IsNotImplemented(e)) { return true; }
        }

        /// <summary>
        /// Evaluates with the invariant culture, so text functions (TEXT, DOLLAR, VALUE...) do not depend on the machine's locale.
        /// Dispose restores the culture. A struct: entering it per cell costs no allocation, and no culture change when it is already invariant.
        /// </summary>
        public static InvariantScope EnterInvariant() => new InvariantScope(CultureInfo.CurrentCulture);

        internal readonly struct InvariantScope : IDisposable
        {
            readonly CultureInfo _restore; // null when the thread already ran with the invariant culture

            internal InvariantScope(CultureInfo current)
            {
                _restore = ReferenceEquals(current, CultureInfo.InvariantCulture) ? null : current;
                if (_restore != null) CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
                EvaluationScope.Enter();
            }

            public void Dispose()
            {
                EvaluationScope.Leave();
                if (_restore != null) CultureInfo.CurrentCulture = _restore;
            }
        }

        static object FromEvaluated(WorkBook book, XL.ICell c, XL.CellValue cv, out Kind kind)
        {
            switch (cv.CellType)
            {
                case XL.CellType.Numeric:
                    if (double.IsNaN(cv.NumberValue) || double.IsInfinity(cv.NumberValue)) { kind = Kind.Error; return "#NUM!"; }
                    if (IsDateValue(book, c, cv.NumberValue) && Serial.TryToDateTime(cv.NumberValue, out var day)) { kind = Kind.Date; return day; }
                    kind = Kind.Number; return cv.NumberValue;
                case XL.CellType.String: kind = Kind.Text; return cv.StringValue;
                case XL.CellType.Boolean: kind = Kind.Bool; return cv.BooleanValue;
                case XL.CellType.Error: kind = Kind.Error; return XL.FormulaError.ForInt(cv.ErrorValue).String;
                default: kind = Kind.Empty; return null;
            }
        }

        static object FromCell(WorkBook book, XL.ICell c, XL.CellType t, out Kind kind)
        {
            switch (t)
            {
                case XL.CellType.Numeric:
                    var serial = c.NumericCellValue; // an XLSX cell keeps the number as text and parses it on every read
                    if (IsDateValue(book, c, serial) && Serial.TryToDateTime(serial, out var day)) { kind = Kind.Date; return day; }
                    kind = Kind.Number; return serial;
                case XL.CellType.String: kind = Kind.Text; return c.StringCellValue;
                case XL.CellType.Boolean: kind = Kind.Bool; return c.BooleanCellValue;
                case XL.CellType.Error: kind = Kind.Error; return XL.FormulaError.ForInt(c.ErrorCellValue).String;
                default: kind = Kind.Empty; return null;
            }
        }

        public static bool IsClearing(object v) => v == null || v is DBNull || (v is string s && string.IsNullOrWhiteSpace(s));

        /// <summary>Stores a non-clearing value. Callers handle <see cref="IsClearing"/> first.</summary>
        public static void Write(WorkBook book, XL.ICell c, object v)
        {
            switch (v)
            {
                case string s: c.SetCellValue(s); return;
                case char ch: c.SetCellValue(ch.ToString()); return;
                case bool b: c.SetCellValue(b); return;
                case DateTime d: WriteDate(book, c, d); return;
                case DateTimeOffset o: WriteDate(book, c, o.DateTime); return;
                case byte _: case sbyte _: case short _: case ushort _: case int _: case uint _:
                case long _: case ulong _: case float _: case double _: case decimal _:
                    WriteNumber(c, Convert.ToDouble(v, CultureInfo.InvariantCulture)); return;
                default: c.SetCellValue(Convert.ToString(v, CultureInfo.InvariantCulture)); return;
            }
        }

        static void WriteNumber(XL.ICell c, double d)
        {
            // Excel cannot store NaN/Infinity (the file would be corrupt); #NUM! is what Excel itself shows.
            if (double.IsNaN(d) || double.IsInfinity(d)) c.SetCellErrorValue((byte)XL.FormulaError.NUM.Code);
            else c.SetCellValue(d);
        }

        public static void WriteDate(WorkBook book, XL.ICell c, DateTime d)
        {
            c.SetCellValue(Serial.FromDateTime(d));
            if (!IsDateFormatted(book, c)) book.Styles.Apply(c, "numfmt:date", s => s.DataFormat = DefaultDateFormat);
        }

        public static string Raw(object v)
        {
            switch (v)
            {
                case null: return "";
                case double d: return d.ToString("R", CultureInfo.InvariantCulture);
                case DateTime dt: return dt.ToString("G", CultureInfo.InvariantCulture);
                case bool b: return b ? "TRUE" : "FALSE";
                default: return Convert.ToString(v, CultureInfo.InvariantCulture);
            }
        }

        public static double ToDouble(object v)
        {
            switch (v)
            {
                case double d: return d;
                case DateTime dt: return Serial.FromDateTime(dt);
                default: return 0;
            }
        }
    }

    /// <summary>What a cell style says about showing a number: the format code, its index, and whether it draws dates.</summary>
    internal sealed class NumberStyle
    {
        NumberStyle(int index, string format, bool isDate) { Index = index; Format = format; IsDate = isDate; }

        public int Index { get; }
        public string Format { get; }
        public bool IsDate { get; }

        /// <summary>Reads it from the cell's style; a cell with no style shows General.</summary>
        public static NumberStyle Of(XL.ICell cell)
        {
            var style = cell.CellStyle;
            if (style == null) return new NumberStyle(0, "General", false);
            var index = style.DataFormat;
            var format = style.GetDataFormatString();
            return new NumberStyle(index, format, XL.DateUtil.IsADateFormat(index, format));
        }
    }

    /// <summary>Generic conversion used by <c>GetValue&lt;T&gt;</c>.</summary>
    internal static class Conv
    {
        public static T To<T>(object v)
        {
            var t = typeof(T);
            if (v == null)
            {
                if (!t.IsValueType || Nullable.GetUnderlyingType(t) != null) return default(T);
                throw new InvalidCastException($"The cell is empty and cannot be converted to {t.Name}.");
            }
            if (v is T already) return already;
            var target = Nullable.GetUnderlyingType(t) ?? t;
            if (target == typeof(string)) return (T)(object)CellValues.Raw(v);
            if (target == typeof(DateTime) && v is double serial) return (T)(object)Serial.ToDateTime(serial);
            if (target == typeof(object)) return (T)v;
            if (target.IsEnum) return (T)Enum.Parse(target, Convert.ToString(v, CultureInfo.InvariantCulture), true);
            if (v is DateTime dt && target != typeof(DateTime)) v = Serial.FromDateTime(dt);
            return (T)Convert.ChangeType(v, target, CultureInfo.InvariantCulture);
        }

        public static int ToInt(double d) => d >= int.MaxValue ? int.MaxValue : d <= int.MinValue ? int.MinValue : (int)d;
        public static long ToLong(double d) => d >= long.MaxValue ? long.MaxValue : d <= long.MinValue ? long.MinValue : (long)d;
    }
}
