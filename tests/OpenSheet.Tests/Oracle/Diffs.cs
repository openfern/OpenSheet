using System;
using System.Collections.Generic;
using System.Globalization;
using Xunit;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>Collects every mismatch so one run reports them all instead of stopping at the first.</summary>
    public sealed class Diffs
    {
        readonly List<string> _list = new List<string>();

        public void Eq(string what, object expected, object actual)
        {
            if (!Same(expected, actual)) _list.Add($"{what}: expected [{Show(expected)}] but Excel says [{Show(actual)}]");
        }

        public void True(string what, bool condition) { if (!condition) _list.Add($"{what}: expected true"); }

        /// <summary>Runs a check; an exception from Excel's side is recorded as a mismatch, not a crash.</summary>
        public void Try(string what, Action check)
        {
            try { check(); }
            catch (Exception e) { _list.Add($"{what}: {e.GetType().Name}: {e.Message.Split('\n')[0]}"); }
        }

        public void Assert()
        {
            if (_list.Count > 0) Xunit.Assert.Fail($"{_list.Count} difference(s) against Excel:\n  " + string.Join("\n  ", _list));
        }

        /// <summary>Excel hands error values over COM as large negative integers; this is their usual text.</summary>
        public static string ExcelErrorText(int code)
        {
            switch (code)
            {
                case -2146826281: return "#DIV/0!";
                case -2146826246: return "#N/A";
                case -2146826259: return "#NAME?";
                case -2146826288: return "#NULL!";
                case -2146826252: return "#NUM!";
                case -2146826265: return "#REF!";
                case -2146826273: return "#VALUE!";
                case -2146826240: return "#SPILL!";
                case -2146826280: return "#CALC!";
                default: return "#ERR" + code;
            }
        }

        static bool Same(object a, object b)
        {
            if (a == null || b == null) return a == null && b == null;
            if (IsNumber(a) && IsNumber(b)) return Math.Abs(Convert.ToDouble(a, CultureInfo.InvariantCulture) - Convert.ToDouble(b, CultureInfo.InvariantCulture)) < 1e-9;
            return string.Equals(Convert.ToString(a, CultureInfo.InvariantCulture), Convert.ToString(b, CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        static bool IsNumber(object o) => o is double || o is float || o is int || o is long || o is short || o is decimal || o is byte || o is uint;

        static string Show(object o) => o == null ? "null" : Convert.ToString(o, CultureInfo.InvariantCulture);
    }
}
