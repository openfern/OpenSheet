using System;
using System.Linq;
using System.Text.RegularExpressions;
using OpenSheet.Internal.Formulas;
using Xunit;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Every function the library adds (or corrects) has to appear in the bank that is compared with Excel, so a new function
    /// cannot ship with nothing but the author's belief that it is right.
    /// </summary>
    public class FormulaBankCoverageTests
    {
        /// <summary>Functions whose answer depends on the machine Excel runs on, so no recording can stand for it.</summary>
        static readonly System.Collections.Generic.Dictionary<string, string> Unverifiable = new System.Collections.Generic.Dictionary<string, string>
        {
            ["JIS"] = "Excel converts to full-width characters only in East-Asian installations and returns the text unchanged elsewhere",
        };

        [Fact]
        public void EveryFunctionTheLibraryAdds_IsInTheBankComparedWithExcel()
        {
            var bank = string.Join("\n", FormulaBank.Formulas);
            var missing = FunctionPack.Names
                .Where(n => !Unverifiable.ContainsKey(n))
                .Where(n => !Regex.IsMatch(bank, @"(?<![A-Z0-9_.])" + Regex.Escape(n) + @"\(", RegexOptions.IgnoreCase))
                .OrderBy(n => n, StringComparer.Ordinal).ToList();
            Assert.True(missing.Count == 0, "Functions registered but never checked against Excel: " + string.Join(", ", missing));
        }
    }
}
