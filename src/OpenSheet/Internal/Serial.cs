using System;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>
    /// Excel's date serial number against .NET's DateTime. A serial between 0 and 1 is a time of day with no date ("18:00"); it maps to
    /// 30 December 1899 plus the time (the same day OLE automation uses as day zero), so reading it and writing it back gives the same number.
    /// </summary>
    internal static class Serial
    {
        static readonly DateTime TimeOnlyDay = new DateTime(1899, 12, 30);

        public static DateTime ToDateTime(double serial) =>
            serial >= 0 && serial < 1 ? TimeOnlyDay.AddMilliseconds(Math.Round(serial * 86400000)) : XL.DateUtil.GetJavaDate(serial);

        /// <summary>False for a serial NPOI cannot turn into a date (negative, not a number, or past what a date can hold), instead of throwing.</summary>
        public static bool TryToDateTime(double serial, out DateTime value)
        {
            try { value = ToDateTime(serial); return true; }
            catch (ArgumentException) { value = default(DateTime); return false; }
        }

        public static double FromDateTime(DateTime value)
        {
            if (value.Date == TimeOnlyDay) return value.TimeOfDay.TotalDays;
            var serial = XL.DateUtil.GetExcelDate(value);
            if (serial < 0) throw new ArgumentOutOfRangeException(nameof(value), value, "Excel cannot hold a date before 1 January 1900 (only a time of day on 30 December 1899).");
            return serial;
        }
    }
}
