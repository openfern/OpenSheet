using System;
using System.Text.RegularExpressions;
using OpenSheet.Internal;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>The hand-written A1 scanner against the regex-and-NPOI implementation it replaced.</summary>
    public class A1Tests
    {
        // ---- the previous implementation, kept as the oracle ----

        static readonly Regex OldCell = new Regex(@"^([A-Za-z]{1,3})([0-9]{1,7})$");

        static string OldStrip(string s)
        {
            s = s.Trim();
            var bang = s.LastIndexOf('!');
            if (bang >= 0) s = s.Substring(bang + 1);
            return s.Replace("$", "");
        }

        static (int row, int col) OldParseCell(string text)
        {
            var m = OldCell.Match(OldStrip(text));
            if (!m.Success) throw new ArgumentException("format");
            var letters = m.Groups[1].Value.ToUpperInvariant();
            var col = NPOI.SS.Util.CellReference.ConvertColStringToIndex(letters);
            if (col > A1.MaxColumn) throw new ArgumentException("column");
            var row = int.Parse(m.Groups[2].Value) - 1;
            if (row < 0 || row > A1.MaxRow) throw new ArgumentException("row");
            return (row, col);
        }

        static RangeAddress OldParse(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("empty");
            var parts = OldStrip(text).Split(':');
            if (parts.Length > 2) throw new ArgumentException("parts");
            var (r1, c1) = OldParseCell(parts[0]);
            var (r2, c2) = parts.Length == 2 ? OldParseCell(parts[1]) : (r1, c1);
            return new RangeAddress(Math.Min(r1, r2), Math.Max(r1, r2), Math.Min(c1, c2), Math.Max(c1, c2));
        }

        // ---- columns ----

        [Fact]
        public void ColumnNames_MatchNpoi_ForEveryColumn()
        {
            for (var i = 0; i <= A1.MaxColumn; i++)
            {
                var name = A1.Column(i);
                Assert.Equal(NPOI.SS.Util.CellReference.ConvertNumToColString(i), name);
                Assert.Equal(i, A1.Column(name));
                Assert.Equal(i, A1.Column("$" + name.ToLowerInvariant()));
            }
        }

        [Theory]
        [InlineData(0, "A")]
        [InlineData(25, "Z")]
        [InlineData(26, "AA")]
        [InlineData(701, "ZZ")]
        [InlineData(702, "AAA")]
        [InlineData(16383, "XFD")]
        public void ColumnNames_Known(int index, string name)
        {
            Assert.Equal(name, A1.Column(index));
            Assert.Equal(index, A1.Column(name));
        }

        [Fact]
        public void ColumnName_IsCached()
        {
            Assert.Same(A1.Column(5), A1.Column(5));
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("$")]
        [InlineData("A1")]
        [InlineData("A B")]
        [InlineData("XFE")]
        [InlineData("AAAA")]
        [InlineData("ZZZZZZZZZZZZZZZZZZZZ")]
        public void ColumnLetters_Invalid_Throw(string letters) =>
            Assert.Throws<ArgumentException>(() => A1.Column(letters));

        [Fact]
        public void ColumnLetters_Null_Throws() => Assert.Throws<ArgumentException>(() => A1.Column((string)null));

        // ---- cells and ranges ----

        [Theory]
        [InlineData("A1", 0, 0)]
        [InlineData("$B$2", 1, 1)]
        [InlineData("Sheet1!C3", 2, 2)]
        [InlineData("'My Sheet'!C3", 2, 2)]
        [InlineData("  d4  ", 3, 3)]
        [InlineData("XFD1048576", 1048575, 16383)]
        [InlineData("a$1", 0, 0)]
        public void ParseCell_Valid(string text, int row, int col)
        {
            A1.ParseCell(text, out var r, out var c);
            Assert.Equal((row, col), (r, c));
        }

        [Theory]
        [InlineData("")]
        [InlineData("A")]
        [InlineData("1")]
        [InlineData("A0")]
        [InlineData("A1048577")]
        [InlineData("XFE1")]
        [InlineData("AAAA1")]
        [InlineData("1A")]
        [InlineData("A1B")]
        [InlineData("A12345678")]
        [InlineData("A 1")]
        [InlineData("A-1")]
        public void ParseCell_Invalid_Throws(string text) =>
            Assert.Throws<ArgumentException>(() => A1.ParseCell(text, out _, out _));

        [Fact]
        public void ParseCell_Null_Throws() => Assert.Throws<ArgumentNullException>(() => A1.ParseCell(null, out _, out _));

        [Theory]
        [InlineData("A1", 0, 0, 0, 0)]
        [InlineData("A1:C9", 0, 8, 0, 2)]
        [InlineData("C9:A1", 0, 8, 0, 2)]
        [InlineData("$A$1:$C$9", 0, 8, 0, 2)]
        [InlineData("Sheet1!A1:C9", 0, 8, 0, 2)]
        [InlineData(" A1 : B2 ", 0, 1, 0, 1)]
        public void Parse_Valid(string text, int r1, int r2, int c1, int c2)
        {
            var a = A1.Parse(text);
            Assert.Equal((r1, r2, c1, c2), (a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn));
        }

        [Theory]
        [InlineData("")]
        [InlineData("  ")]
        [InlineData("A1:")]
        [InlineData(":A1")]
        [InlineData("A1:B2:C3")]
        [InlineData("A:B")]
        public void Parse_Invalid_Throws(string text) => Assert.Throws<ArgumentException>(() => A1.Parse(text));

        [Fact]
        public void Cell_BuildsAddress()
        {
            Assert.Equal("B3", A1.Cell(2, 1));
            Assert.Equal("XFD1048576", A1.Cell(A1.MaxRow, A1.MaxColumn));
        }

        [Theory]
        [InlineData("A:C", true, "A", "C")]
        [InlineData("$a:$bc", true, "a", "bc")]
        [InlineData("1:5", false, "1", "5")]
        [InlineData("$2:$9999999", false, "2", "9999999")]
        public void TryWhole_Matches(string text, bool letters, string first, string second)
        {
            Assert.True(A1.TryWhole(text, letters, out var a, out var b));
            Assert.Equal((first, second), (a, b));
        }

        [Theory]
        [InlineData("A1:C3", true)]
        [InlineData("A:", true)]
        [InlineData("A:B:C", true)]
        [InlineData("AAAA:B", true)]
        [InlineData("1:5", true)]
        [InlineData("A:C", false)]
        [InlineData("12345678:1", false)]
        [InlineData("A1", true)]
        public void TryWhole_Rejects(string text, bool letters) => Assert.False(A1.TryWhole(text, letters, out _, out _));

        // ---- differential: random text through old and new ----

        [Fact]
        public void Parse_AgreesWithPreviousImplementation_OnRandomText()
        {
            const string alphabet = "AaBzXFDx019$:! '_-";
            var rnd = new Random(20260101);
            for (var n = 0; n < 60000; n++)
            {
                var chars = new char[rnd.Next(1, 14)];
                for (var i = 0; i < chars.Length; i++) chars[i] = alphabet[rnd.Next(alphabet.Length)];
                var text = new string(chars);

                RangeAddress expected = null;
                try { expected = OldParse(text); } catch (ArgumentException) { }

                RangeAddress actual = null;
                try { actual = A1.Parse(text); } catch (ArgumentException) { }

                Assert.True(expected == null ? actual == null : expected.Equals(actual), $"'{text}': expected {expected?.Location ?? "error"}, got {actual?.Location ?? "error"}");
            }
        }
    }
}
