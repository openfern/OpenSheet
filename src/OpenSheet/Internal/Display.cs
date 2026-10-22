using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>The text a cell shows, formatting applied, the way Excel draws it (invariant culture).</summary>
    internal static class Display
    {
        // Largest serial Excel can show as a date (31 Dec 9999).
        const double MaxDateSerial = 2958465.99999999;

        /// <summary>The display text of a cell that may not exist; when drawing it fails, the raw value.</summary>
        public static string Formatted(WorkBook book, XL.ICell cell)
        {
            if (cell == null) return "";
            try { return Text(book, cell) ?? ""; }
            catch (Exception) { return CellValues.Raw(CellValues.Read(book, cell, out _)); }
        }

        public static string Text(WorkBook book, XL.ICell cell)
        {
            // text, booleans and blanks do not depend on the culture, a format or a formula: skip the evaluator and the culture switch
            switch (cell.CellType)
            {
                case XL.CellType.String: return cell.StringCellValue;
                case XL.CellType.Boolean: return cell.BooleanCellValue ? "TRUE" : "FALSE";
                case XL.CellType.Blank: return "";
            }
            using (CellValues.EnterInvariant())
            {
                if (cell.CellType == XL.CellType.Formula)
                {
                    XL.CellValue evaluated;
                    try { evaluated = book.Evaluator.Evaluate(cell); }
                    catch (Exception e) when (CellValues.IsNotImplemented(e)) { return "#NAME?"; }
                    if (evaluated != null && evaluated.CellType == XL.CellType.Numeric) return Number(book, cell, evaluated.NumberValue);
                    return book.Formatter.FormatCellValue(cell, book.Evaluator);
                }
                if (cell.CellType == XL.CellType.Numeric) return Number(book, cell, cell.NumericCellValue);
                return book.Formatter.FormatCellValue(cell, book.Evaluator);
            }
        }

        static string Number(WorkBook book, XL.ICell cell, double value)
        {
            var numberStyle = book.NumberStyleOf(cell);
            var format = numberStyle.Format;
            var index = numberStyle.Index;
            if (double.IsNaN(value) || double.IsInfinity(value)) return "#NUM!";

            if (numberStyle.IsDate)
            {
                // Excel shows #### when a date or time cannot be drawn (negative or past the year 9999).
                if (value < 0 || value > MaxDateSerial) return "########";
                if (DateFormat.TryFormat(value, format, out var drawn)) return drawn;
                return book.Formatter.FormatRawCellContents(value, index, format);
            }

            // "General" and the text format "@" show a number as Excel's General does.
            if (format == "General" || format == "@") return General(value);
            return Numeric(book, value, index, format);
        }

        // NPOI reads the letters inside "quotes" as format codes ("EUR" turns into scientific notation), drops the padding "_x"
        // and keeps the fill "*x", which Excel spreads over the column width. Literals are swapped for private-use characters
        // that NPOI copies through untouched, then put back.
        static string Numeric(WorkBook book, double value, int index, string format)
        {
            var literals = new System.Collections.Generic.List<string>();
            var masked = MaskLiterals(format, literals);
            // "General" with text around it, e.g. General"!": the number part is General, the rest is literal
            string text;
            if (masked.IndexOf("General", StringComparison.OrdinalIgnoreCase) >= 0 && masked.IndexOf(';') < 0)
                text = Regex.Replace(masked, "General", General(value), RegexOptions.IgnoreCase);
            else
                text = book.Formatter.FormatRawCellContents(value, index, masked);
            for (var i = 0; i < literals.Count; i++) text = text.Replace(((char)(0xE000 + i)).ToString(), literals[i]);
            // a negative number that rounds to zero shows without its sign
            if (value < 0 && text.StartsWith("-", StringComparison.Ordinal) && text.IndexOfAny(new[] { '1', '2', '3', '4', '5', '6', '7', '8', '9' }) < 0 && masked.IndexOf(';') < 0)
                text = text.Substring(1);
            return text;
        }

        static string MaskLiterals(string format, System.Collections.Generic.List<string> literals)
        {
            if (format.IndexOf('"') < 0 && format.IndexOf('_') < 0 && format.IndexOf('*') < 0 && format.IndexOf('\\') < 0) return format;
            string Keep(string text) { literals.Add(text); return ((char)(0xE000 + literals.Count - 1)).ToString(); }
            var sb = new StringBuilder(format.Length + 8);
            for (var i = 0; i < format.Length; i++)
            {
                var ch = format[i];
                if (ch == '"')
                {
                    var end = format.IndexOf('"', i + 1);
                    if (end < 0) { sb.Append(format, i, format.Length - i); break; }
                    sb.Append(Keep(format.Substring(i + 1, end - i - 1)));
                    i = end;
                }
                else if (ch == '\\' && i + 1 < format.Length) { sb.Append(Keep(format[++i].ToString())); }
                else if (ch == '_' && i + 1 < format.Length) { sb.Append(Keep(" ")); i++; }
                else if (ch == '*' && i + 1 < format.Length) { i++; }
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Excel's General format: at most 11 characters (not counting the sign), fixed notation or scientific, whichever keeps
        /// more of the number's digits (fixed wins a tie). 0.5 → "0.5", 1/3 → "0.333333333", 123456789012 → "1.23457E+11".
        /// </summary>
        internal static string General(double v) =>
            v > -1e11 && v < 1e11 && v == Math.Floor(v) ? ((long)v).ToString(CultureInfo.InvariantCulture) // whole numbers up to 11 digits are always fixed
            : GeneralFull(v);

        /// <summary>The whole algorithm, with no shortcut. Kept apart so a test can hold the shortcut in <see cref="General"/> to it.</summary>
        internal static string GeneralFull(double v)
        {
            if (v == 0) return "0";
            var sign = v < 0 ? "-" : "";
            var a = Math.Abs(v);
            var fixedText = FixedGeneral(a, out var fixedScore);
            var sciText = ScientificGeneral(a, out var sciScore);
            if (fixedText != null && (sciText == null || fixedScore >= sciScore)) return sign + fixedText;
            return sign + sciText;
        }

        static readonly string[] Fixed = { "F0", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10" }; // decimals never exceed 10

        static int Exponent(double a)
        {
            var e = (int)Math.Floor(Math.Log10(a));
            // Log10 can be off by one next to a power of ten; correct it with the exact scientific text.
            var check = a.ToString("E14", CultureInfo.InvariantCulture);
            var p = check.IndexOf('E');
            return p > 0 ? int.Parse(check.Substring(p + 1), CultureInfo.InvariantCulture) : e;
        }

        static string FixedGeneral(double a, out int score)
        {
            score = 0;
            var e = Exponent(a);
            var intDigits = e >= 0 ? e + 1 : 1;
            if (intDigits > 11) return null;
            var decimals = Math.Max(0, 10 - intDigits);
            if (intDigits == 11) decimals = 0;
            var rounded = Math.Round(a, Math.Min(decimals, 15), MidpointRounding.AwayFromZero);
            if (rounded == 0) return null;
            var s = rounded.ToString(Fixed[decimals], CultureInfo.InvariantCulture);
            if (s.IndexOf('.') >= 0) s = s.TrimEnd('0').TrimEnd('.');
            var significant = s.Replace(".", "").TrimStart('0').Length;
            if (significant > 15) return null;
            score = significant;
            return s;
        }

        static string ScientificGeneral(double a, out int score)
        {
            score = 0;
            var e = Exponent(a);
            for (var attempt = 0; attempt < 2; attempt++)
            {
                var expText = (e < 0 ? "-" : "+") + Math.Abs(e).ToString(Math.Abs(e) < 100 ? "00" : "000", CultureInfo.InvariantCulture);
                var decimals = Math.Max(0, 11 - (1 + expText.Length) - 2); // 11 characters: "d", ".", the decimals, "E" and the exponent
                var mantissa = Math.Round(a / Math.Pow(10, e), decimals, MidpointRounding.AwayFromZero);
                if (mantissa >= 10) { e++; continue; } // 9.99999 rounded up to 10: move to the next exponent
                var m = mantissa.ToString(Fixed[decimals], CultureInfo.InvariantCulture);
                if (m.IndexOf('.') >= 0) m = m.TrimEnd('0').TrimEnd('.');
                score = m.Replace(".", "").Length;
                return m + "E" + expText;
            }
            return null;
        }
    }
}
