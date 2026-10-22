using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenSheet.Internal;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>The CSV reader against the version that built a list of lists first; same rules, one row alive at a time.</summary>
    public class CsvParserTests
    {
        static List<List<string>> Reference(string text, string delimiter, bool ignoreQuotes)
        {
            var rows = new List<List<string>>();
            var row = new List<string>(); var sb = new StringBuilder();
            var inQuotes = false; var fieldStarted = false;
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                if (inQuotes)
                {
                    if (ch == '"' && i + 1 < text.Length && text[i + 1] == '"') { sb.Append('"'); i++; }
                    else if (ch == '"') inQuotes = false;
                    else sb.Append(ch);
                }
                else if (ch == '"' && !ignoreQuotes && sb.Length == 0) { inQuotes = true; fieldStarted = true; }
                else if (string.CompareOrdinal(text, i, delimiter, 0, delimiter.Length) == 0)
                {
                    row.Add(sb.ToString()); sb.Clear(); fieldStarted = false; i += delimiter.Length - 1;
                }
                else if (ch == '\r' || ch == '\n')
                {
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    if (row.Count > 0 || sb.Length > 0 || fieldStarted) { row.Add(sb.ToString()); rows.Add(row); }
                    row = new List<string>(); sb.Clear(); fieldStarted = false;
                }
                else sb.Append(ch);
            }
            if (row.Count > 0 || sb.Length > 0 || fieldStarted) { row.Add(sb.ToString()); rows.Add(row); }
            return rows;
        }

        static List<List<string>> Parse(string text, string delimiter, bool ignoreQuotes)
        {
            var rows = new List<List<string>>();
            TextImport.ParseCsv(text, delimiter, ignoreQuotes, row => rows.Add(new List<string>(row))); // the list is reused, so copy it
            return rows;
        }

        [Theory]
        [InlineData(",", false)]
        [InlineData(",", true)]
        [InlineData(";;", false)]
        [InlineData("\t", false)]
        [InlineData("ab", true)]
        public void Parse_AgreesWithThePreviousParser_OnRandomText(string delimiter, bool ignoreQuotes)
        {
            const string alphabet = "ab,;\t\"\"\r\n x";
            var rnd = new Random(31337);
            for (var n = 0; n < 20000; n++)
            {
                var chars = new char[rnd.Next(0, 40)];
                for (var i = 0; i < chars.Length; i++) chars[i] = alphabet[rnd.Next(alphabet.Length)];
                var text = new string(chars);
                var expected = Reference(text, delimiter, ignoreQuotes);
                var actual = Parse(text, delimiter, ignoreQuotes);
                Assert.True(expected.Count == actual.Count && expected.Zip(actual, (a, b) => a.SequenceEqual(b)).All(x => x),
                    "text " + text.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t") + " with delimiter '" + delimiter + "'");
            }
        }

        [Fact]
        public void Parse_KeepsEmptyQuotedFields_AndTheLastRowWithoutALineEnd()
        {
            var rows = Parse("a,\"\",c\r\n\"\"\r\nx,y", ",", false);
            Assert.Equal(3, rows.Count);
            Assert.Equal(new[] { "a", "", "c" }, rows[0]);
            Assert.Equal(new[] { "" }, rows[1]);
            Assert.Equal(new[] { "x", "y" }, rows[2]);
        }

        [Theory]
        [InlineData("a,b;c\r\n", ",")]
        [InlineData("\r\n\r\n  \r\na;b;c", ";")]
        [InlineData("x|y\n", "|")]
        [InlineData("\"a,b\";c;d\n", ";")]
        [InlineData("", ",")]
        [InlineData("   \n\t\n", ",")]
        public void GuessDelimiter_LooksOnlyAtTheFirstLineWithContent(string text, string expected) =>
            Assert.Equal(expected, TextImport.GuessDelimiter(text));
    }
}
