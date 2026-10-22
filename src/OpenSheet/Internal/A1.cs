using System;
using Cysharp.Text;

namespace OpenSheet.Internal
{
    /// <summary>A1-notation helpers. All indexes are zero-based. Parsing scans the text once and allocates nothing.</summary>
    internal static class A1
    {
        public const int MaxRow = 1048575;
        public const int MaxColumn = 16383;

        // "A" .. "AMJ" (LibreOffice's last column). Filled on first use; two threads racing store equal strings, so no lock.
        static readonly string[] ColumnNames = new string[1024];

        public static string Column(int index)
        {
            if ((uint)index >= (uint)ColumnNames.Length) return index >= 0 && index <= MaxColumn ? Letters(index) : NPOI.SS.Util.CellReference.ConvertNumToColString(index);
            return ColumnNames[index] ?? (ColumnNames[index] = Letters(index));
        }

        /// <summary>Bijective base 26: 0 is "A", 25 is "Z", 26 is "AA", 16383 is "XFD".</summary>
        static string Letters(int index)
        {
            var n = index + 1;
            var chars = new char[n <= 26 ? 1 : n <= 702 ? 2 : 3];
            for (var i = chars.Length - 1; i >= 0; i--) { n--; chars[i] = (char)('A' + n % 26); n /= 26; }
            return new string(chars);
        }

        /// <summary>Column letters ("B", "$b", " AA ") to a zero-based index.</summary>
        public static int Column(string letters)
        {
            if (string.IsNullOrWhiteSpace(letters)) throw new ArgumentException("Column name is empty.", nameof(letters));
            return Column(letters, 0, letters.Length);
        }

        /// <summary>The column of a cell reference such as "AB12": its leading letters.</summary>
        public static int ColumnOf(string cellRef)
        {
            var n = 0;
            while (n < cellRef.Length && char.IsLetter(cellRef[n])) n++;
            return Column(cellRef, 0, n);
        }

        static int Column(string text, int from, int to)
        {
            int lo = from, hi = to;
            Trim(text, ref lo, ref hi);
            var n = 0; var seen = 0; var over = false;
            for (var i = lo; i < hi; i++)
            {
                var ch = text[i];
                if (ch == '$') continue;
                if (ch >= 'a' && ch <= 'z') ch = (char)(ch - 32);
                if (ch < 'A' || ch > 'Z') throw new ArgumentException($"Invalid column name '{text.Substring(from, to - from)}'.", "letters");
                seen++;
                if (n <= MaxColumn) n = n * 26 + (ch - 'A' + 1); else over = true; // keep validating the rest; no int overflow on long input
            }
            if (seen == 0) throw new ArgumentException($"Invalid column name '{text.Substring(from, to - from)}'.", "letters");
            if (over || n - 1 > MaxColumn) throw new ArgumentException($"Column '{text.Substring(from, to - from)}' is out of range.", "letters");
            return n - 1;
        }

        public static string Cell(int row, int col) => ZString.Concat(Column(col), row + 1);

        public static void ParseCell(string text, out int row, out int col)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            Scan(text, 0, text.Length, false, out row, out col);
        }

        /// <summary>Parses "A1", "A1:C9", "$A$1:$C$9" or "Sheet1!A1:C9". Corners are normalized.</summary>
        public static RangeAddress Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Range address is empty.", nameof(text));
            int lo = 0, hi = text.Length;
            Trim(text, ref lo, ref hi);
            var bang = text.LastIndexOf('!', hi - 1, hi - lo);
            if (bang >= 0) lo = bang + 1;
            var colon = text.IndexOf(':', lo, hi - lo);
            if (colon >= 0 && text.IndexOf(':', colon + 1, hi - colon - 1) >= 0) throw new ArgumentException($"Invalid range address '{text}'.", nameof(text));
            if (colon < 0)
            {
                Scan(text, lo, hi, true, out var row, out var col);
                return new RangeAddress(row, row, col, col);
            }
            Scan(text, lo, colon, true, out var r1, out var c1);
            Scan(text, colon + 1, hi, true, out var r2, out var c2);
            return new RangeAddress(Math.Min(r1, r2), Math.Max(r1, r2), Math.Min(c1, c2), Math.Max(c1, c2));
        }

        /// <summary>
        /// "A:C" or "$A:$C" (<paramref name="letters"/>), "2:5" or "$2:$5" (digits): a whole-column or whole-row range. Returns both ends without '$'.
        /// </summary>
        public static bool TryWhole(string text, bool letters, out string first, out string second)
        {
            first = second = null;
            var colon = text.IndexOf(':');
            if (colon < 0 || colon != text.LastIndexOf(':')) return false;
            return End(text, 0, colon, letters, out first) && End(text, colon + 1, text.Length, letters, out second);
        }

        static bool End(string text, int lo, int hi, bool letters, out string end)
        {
            end = null;
            if (lo < hi && text[lo] == '$') lo++;
            if (hi - lo < 1 || hi - lo > (letters ? 3 : 7)) return false;
            for (var i = lo; i < hi; i++)
            {
                var ch = text[i];
                if (letters ? !(ch >= 'A' && ch <= 'Z' || ch >= 'a' && ch <= 'z') : !(ch >= '0' && ch <= '9')) return false;
            }
            end = text.Substring(lo, hi - lo);
            return true;
        }

        static void Trim(string s, ref int lo, ref int hi)
        {
            while (lo < hi && char.IsWhiteSpace(s[lo])) lo++;
            while (hi > lo && char.IsWhiteSpace(s[hi - 1])) hi--;
        }

        /// <summary>
        /// One cell in text[lo..hi): optional sheet prefix, '$' anywhere, 1-3 letters then 1-7 digits. <paramref name="part"/> is a piece of a range
        /// address, which was stripped of '$' as a whole and is stripped again here, so white space left at its edges by that goes too ("$ A1" is A1).
        /// </summary>
        static void Scan(string text, int lo, int hi, bool part, out int row, out int col)
        {
            Trim(text, ref lo, ref hi);
            if (part)
            {
                while (lo < hi && (text[lo] == '$' || char.IsWhiteSpace(text[lo]))) lo++;
                while (hi > lo && (text[hi - 1] == '$' || char.IsWhiteSpace(text[hi - 1]))) hi--;
            }
            else
            {
                var bang = hi > lo ? text.LastIndexOf('!', hi - 1, hi - lo) : -1;
                if (bang >= 0) lo = bang + 1;
            }

            var letters = 0; var column = 0; var digits = 0; var number = 0;
            for (var i = lo; i < hi; i++)
            {
                var ch = text[i];
                if (ch == '$') continue;
                if (digits == 0 && (ch >= 'A' && ch <= 'Z' || ch >= 'a' && ch <= 'z'))
                {
                    if (++letters > 3) throw Invalid(text, lo, hi);
                    column = column * 26 + ((ch | 0x20) - 'a' + 1);
                }
                else if (letters > 0 && ch >= '0' && ch <= '9')
                {
                    if (++digits > 7) throw Invalid(text, lo, hi);
                    number = number * 10 + (ch - '0');
                }
                else throw Invalid(text, lo, hi);
            }
            if (letters == 0 || digits == 0) throw Invalid(text, lo, hi);
            if (column - 1 > MaxColumn) throw new ArgumentException($"Column '{LettersOf(text, lo, hi)}' is out of range.", "letters");
            if (number < 1 || number - 1 > MaxRow) throw new ArgumentException($"Row in '{text.Substring(lo, hi - lo).Replace("$", "")}' is out of range.", nameof(text));
            col = column - 1;
            row = number - 1;
        }

        static ArgumentException Invalid(string text, int lo, int hi) => new ArgumentException($"Invalid cell address '{text.Substring(lo, hi - lo).Replace("$", "")}'.", nameof(text));

        static string LettersOf(string text, int lo, int hi)
        {
            var end = lo;
            while (end < hi && (text[end] == '$' || !(text[end] >= '0' && text[end] <= '9'))) end++;
            return text.Substring(lo, end - lo).Replace("$", "");
        }
    }
}
