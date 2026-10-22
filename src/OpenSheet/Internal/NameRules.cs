using System;
using System.Text.RegularExpressions;

namespace OpenSheet.Internal
{
    /// <summary>
    /// What Excel accepts as the name of a table or a defined name. A name that breaks these rules makes Excel report a damaged
    /// file, so it is refused here instead of being written.
    /// </summary>
    internal static class NameRules
    {
        static readonly Regex Letters = new Regex(@"^[\p{L}_\\][\p{L}\p{N}_.\\]*$", RegexOptions.CultureInvariant | RegexOptions.Compiled);
        static readonly Regex CellLike = new Regex(@"^\$?([A-Za-z]{1,3})\$?([0-9]+)$", RegexOptions.CultureInvariant | RegexOptions.Compiled);
        static readonly Regex R1C1 = new Regex(@"^[Rr](\[?-?[0-9]*\]?)?([Cc](\[?-?[0-9]*\]?)?)?$|^[Cc](\[?-?[0-9]*\]?)?$", RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static void Validate(string name, string paramName, string what)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException($"The {what} name is empty.", paramName);
            if (name.Length > 255) throw new ArgumentException($"The {what} name '{Short(name)}' is longer than 255 characters.", paramName);
            if (!Letters.IsMatch(name))
                throw new ArgumentException($"'{name}' is not a valid {what} name: it must start with a letter or underscore and contain only letters, digits, underscores and periods (no spaces).", paramName);
            if (LooksLikeACell(name))
                throw new ArgumentException($"'{name}' is not a valid {what} name: Excel reads it as a cell address.", paramName);
        }

        static bool LooksLikeACell(string name)
        {
            var m = CellLike.Match(name);
            if (m.Success)
            {
                var col = 0;
                foreach (var ch in m.Groups[1].Value.ToUpperInvariant()) col = col * 26 + (ch - 'A' + 1);
                if (col <= 16384 && long.TryParse(m.Groups[2].Value, out var row) && row >= 1 && row <= 1048576) return true;
            }
            return R1C1.IsMatch(name);
        }

        static string Short(string s) => s.Substring(0, 40) + "...";
    }
}
