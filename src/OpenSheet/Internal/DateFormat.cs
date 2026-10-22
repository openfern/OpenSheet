using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using Cysharp.Text;

namespace OpenSheet.Internal
{
    /// <summary>
    /// Draws a date/time serial with an Excel format code (yyyy-mm-dd, h:mm AM/PM, [h]:mm:ss, ...), English names, the way Excel does:
    /// the clock is rounded to the precision shown, day 0 is "1900-01-00", and 29 Feb 1900 exists.
    /// </summary>
    internal static class DateFormat
    {
        static readonly string[] Months = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        static readonly string[] Days = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };

        enum Kind { Literal, Year, Month, Day, Hour, Minute, Second, Fraction, AmPm, Elapsed }

        struct Token
        {
            public Kind Kind;
            public string Text;   // literal text, or the AM/PM designators as written ("AM/PM", "a/p", ...)
            public int Count;     // how many letters (mm = 2), or digits of the fraction
            public char Unit;     // for Elapsed: h, m or s
        }

        // A sheet reuses a handful of format codes for thousands of cells, so the parsed form is kept. Bounded: codes come from the file.
        static readonly ConcurrentDictionary<string, List<Token>> Parsed = new ConcurrentDictionary<string, List<Token>>(StringComparer.Ordinal);
        static readonly List<Token> NotADate = new List<Token>();
        const int MaxParsed = 512;

        static List<Token> Tokens(string code)
        {
            if (Parsed.TryGetValue(code, out var tokens)) return tokens;
            if (Parsed.Count >= MaxParsed) Parsed.Clear();
            return Parsed.GetOrAdd(code, c => Parse(FirstSection(c)) ?? NotADate);
        }

        /// <summary>The first section of the code only: dates are never negative here, so the other sections do not apply.</summary>
        public static bool TryFormat(double serial, string code, out string text)
        {
            text = null;
            if (serial < 0 || serial >= 2958466) return false;
            var tokens = Tokens(code);
            if (tokens.Count == 0) return false;

            var fraction = 0;
            foreach (var t in tokens) if (t.Kind == Kind.Fraction) fraction = Math.Max(fraction, t.Count);
            var twelve = tokens.Exists(t => t.Kind == Kind.AmPm);

            // round the instant to what is displayed
            var unit = fraction == 0 ? 1.0 : Math.Pow(10, -fraction);          // seconds
            var seconds = Math.Round(serial * 86400.0 / unit, MidpointRounding.AwayFromZero) * unit;
            var day = Math.Floor(seconds / 86400.0 + 1e-9);
            var secOfDay = seconds - day * 86400.0;
            if (secOfDay < 0) secOfDay = 0;

            GetDate((int)day, out var year, out var month, out var dom, out var weekday);
            var hour = (int)(secOfDay / 3600);
            var minute = (int)(secOfDay % 3600 / 60);
            var second = (int)(secOfDay % 60);
            var frac = secOfDay - Math.Floor(secOfDay);

            var totalSeconds = Math.Floor(seconds + 1e-9);
            var sb = ZString.CreateStringBuilder(); // a plain local, not a using variable: the Pad helper takes it by ref
            try
            {
                foreach (var t in tokens)
                {
                    switch (t.Kind)
                    {
                        case Kind.Literal: sb.Append(t.Text); break;
                        case Kind.Year: Pad(ref sb, t.Count <= 2 ? year % 100 : year, t.Count <= 2 ? 2 : 4); break;
                        case Kind.Month:
                            if (t.Count <= 2) Pad(ref sb, month, t.Count);
                            else if (t.Count == 3) sb.Append(month == 0 ? "" : Months[month - 1].Substring(0, 3));
                            else if (t.Count == 4) sb.Append(month == 0 ? "" : Months[month - 1]);
                            else sb.Append(month == 0 ? "" : Months[month - 1].Substring(0, 1));
                            break;
                        case Kind.Day:
                            if (t.Count <= 2) Pad(ref sb, dom, t.Count);
                            else if (t.Count == 3) sb.Append(Days[weekday].Substring(0, 3));
                            else sb.Append(Days[weekday]);
                            break;
                        case Kind.Hour: Pad(ref sb, twelve ? (hour % 12 == 0 ? 12 : hour % 12) : hour, t.Count); break;
                        case Kind.Minute: Pad(ref sb, minute, t.Count); break;
                        case Kind.Second: Pad(ref sb, second, t.Count); break;
                        case Kind.Fraction:
                            sb.Append('.');
                            Pad(ref sb, (int)Math.Round(frac * Math.Pow(10, t.Count)), t.Count);
                            break;
                        case Kind.AmPm: sb.Append(AmPm(t.Text, hour >= 12)); break;
                        case Kind.Elapsed:
                            {
                                var total = t.Unit == 'h' ? Math.Floor(totalSeconds / 3600) : t.Unit == 'm' ? Math.Floor(totalSeconds / 60) : totalSeconds;
                                Pad(ref sb, (long)total, Math.Max(1, t.Count));
                                break;
                            }
                    }
                }
                text = sb.ToString();
            }
            finally { sb.Dispose(); }
            return true;
        }

        /// <summary>The number with at least <paramref name="width"/> digits, zero padded (the "00" and "0000" of a format code).</summary>
        static void Pad(ref Utf16ValueStringBuilder sb, long value, int width)
        {
            for (var shown = Digits(value); shown < width; shown++) sb.Append('0');
            sb.Append(value);
        }

        static int Digits(long v)
        {
            var n = 1;
            for (v = v < 0 ? -v : v; v >= 10; v /= 10) n++;
            return n;
        }

        static string AmPm(string written, bool pm)
        {
            var upper = char.IsUpper(written[0]);
            var full = written.Length > 1 && written.Substring(0, 2).ToUpperInvariant() == "AM";
            var s = full ? (pm ? "PM" : "AM") : (pm ? "P" : "A");
            return upper ? s : s.ToLowerInvariant();
        }

        // Excel's calendar: serial 60 is the 29th of February 1900, which never existed, and serial 0 is the 0th of January 1900.
        static void GetDate(int serial, out int year, out int month, out int day, out int weekday)
        {
            weekday = (int)(((serial - 1) % 7 + 7) % 7);
            if (serial < 61)
            {
                year = 1900;
                if (serial <= 31) { month = 1; day = serial; }
                else { month = 2; day = serial - 31; }
                return;
            }
            var d = new DateTime(1899, 12, 30).AddDays(serial);
            year = d.Year; month = d.Month; day = d.Day;
        }

        static string FirstSection(string code)
        {
            var inQuote = false;
            for (var i = 0; i < code.Length; i++)
            {
                if (code[i] == '"') inQuote = !inQuote;
                else if (code[i] == '\\') i++;
                else if (code[i] == ';' && !inQuote) return code.Substring(0, i);
            }
            return code;
        }

        static List<Token> Parse(string code)
        {
            var list = new List<Token>();
            var i = 0;
            Token? lastTime = null; // the previous h or [h] decides that the next m is minutes
            while (i < code.Length)
            {
                var ch = code[i];
                if (ch == '"')
                {
                    var end = code.IndexOf('"', i + 1);
                    if (end < 0) end = code.Length;
                    list.Add(Lit(code.Substring(i + 1, Math.Max(0, end - i - 1))));
                    i = Math.Min(code.Length, end + 1);
                    continue;
                }
                if (ch == '\\') { if (i + 1 < code.Length) list.Add(Lit(code[i + 1].ToString())); i += 2; continue; }
                if (ch == '_') { list.Add(Lit(" ")); i += 2; continue; }
                if (ch == '*') { i += 2; continue; }
                if (ch == '[')
                {
                    var end = code.IndexOf(']', i);
                    if (end < 0) return null;
                    var inner = code.Substring(i + 1, end - i - 1);
                    var lower = inner.ToLowerInvariant();
                    if (lower.Length > 0 && (lower[0] == 'h' || lower[0] == 'm' || lower[0] == 's') && lower.Trim(lower[0]).Length == 0)
                    {
                        var t = new Token { Kind = Kind.Elapsed, Unit = lower[0], Count = lower.Length };
                        list.Add(t);
                        lastTime = t.Unit == 'h' ? t : lastTime;
                    }
                    else if (inner.StartsWith("$", StringComparison.Ordinal))
                    {
                        var dash = inner.IndexOf('-');
                        var symbol = dash < 0 ? inner.Substring(1) : inner.Substring(1, dash - 1);
                        if (symbol.Length > 0) list.Add(Lit(symbol));
                    }
                    // colors, conditions and other brackets do not draw anything
                    i = end + 1;
                    continue;
                }
                var lc = char.ToLowerInvariant(ch);
                if (lc == 'a' && i + 5 <= code.Length && string.Compare(code, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
                { list.Add(new Token { Kind = Kind.AmPm, Text = code.Substring(i, 5) }); i += 5; continue; }
                if (lc == 'a' && i + 3 <= code.Length && string.Compare(code, i, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0)
                { list.Add(new Token { Kind = Kind.AmPm, Text = code.Substring(i, 3) }); i += 3; continue; }
                if (lc == 'y' || lc == 'd' || lc == 'h' || lc == 's' || lc == 'm')
                {
                    var n = 1;
                    while (i + n < code.Length && char.ToLowerInvariant(code[i + n]) == lc) n++;
                    i += n;
                    switch (lc)
                    {
                        case 'y': list.Add(new Token { Kind = Kind.Year, Count = n >= 3 ? 4 : 2 }); lastTime = null; break;
                        case 'd': list.Add(new Token { Kind = Kind.Day, Count = Math.Min(n, 4) }); lastTime = null; break;
                        case 'h': { var t = new Token { Kind = Kind.Hour, Count = Math.Min(n, 2) }; list.Add(t); lastTime = t; break; }
                        case 's':
                            list.Add(new Token { Kind = Kind.Second, Count = Math.Min(n, 2) });
                            if (i + 1 < code.Length && code[i] == '.' && code[i + 1] == '0')
                            {
                                var z = 0; // digits of the fraction: ".0" is one, ".000" is three
                                while (i + 1 + z < code.Length && code[i + 1 + z] == '0' && z < 3) z++;
                                list.Add(new Token { Kind = Kind.Fraction, Count = z });
                                i += 1 + z;
                            }
                            lastTime = null;
                            break;
                        default:
                            {
                                // minutes right after hours, or right before seconds; otherwise months
                                var asMinutes = lastTime != null || NextIsSeconds(code, i);
                                list.Add(asMinutes ? new Token { Kind = Kind.Minute, Count = Math.Min(n, 2) } : new Token { Kind = Kind.Month, Count = Math.Min(n, 5) });
                                lastTime = null;
                                break;
                            }
                    }
                    continue;
                }
                if (char.IsDigit(ch) || ch == '#' || ch == '?' || ch == 'E' || ch == 'e' || ch == '@') return null; // a number or text code, not a date code
                list.Add(Lit(ch.ToString()));
                i++;
            }
            return list;
        }

        static bool NextIsSeconds(string code, int from)
        {
            for (var i = from; i < code.Length; i++)
            {
                var c = char.ToLowerInvariant(code[i]);
                if (c == 's') return true;
                if (c == '"' || c == '\\' || c == '[') return false;
                if (c == 'h' || c == 'd' || c == 'y' || c == 'm') return false;
            }
            return false;
        }

        static Token Lit(string s) => new Token { Kind = Kind.Literal, Text = s };
    }
}
