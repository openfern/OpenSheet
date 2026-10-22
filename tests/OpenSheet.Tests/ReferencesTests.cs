using OpenSheet.Internal.Formulas;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>How references inside formula text follow copies and structural edits (the Excel oracle checks the same on a real workbook).</summary>
    public class ReferencesTests
    {
        [Theory]
        [InlineData("A1*2", 0, 2, "C1*2")]
        [InlineData("$A1+B$2", 3, 1, "$A4+C$2")]
        [InlineData("$A$1+B2", 5, 5, "$A$1+G7")]
        [InlineData("SUM(A1:B5)", 1, 1, "SUM(B2:C6)")]
        [InlineData("SUM($A$1:B5)", 1, 1, "SUM($A$1:C6)")]
        [InlineData("S!A1+'My Sheet'!B2", 1, 0, "S!A2+'My Sheet'!B3")]
        [InlineData("\"A1 stays\"&A1", 1, 0, "\"A1 stays\"&A2")]
        [InlineData("LOG10(A1)+ATAN2(B1,C1)", 0, 1, "LOG10(B1)+ATAN2(C1,D1)")]
        [InlineData("SUM(A:B)+SUM(1:2)", 1, 1, "SUM(B:C)+SUM(2:3)")]
        [InlineData("SUM($A:$B)+SUM($1:$2)", 4, 4, "SUM($A:$B)+SUM($1:$2)")]
        [InlineData("Total*A1", 0, 1, "Total*B1")]
        [InlineData("A1", -1, 0, "#REF!")]
        [InlineData("S!A1", 0, -1, "S!#REF!")]
        [InlineData("SUM(A1:A3)", -1, 0, "SUM(#REF!)")]
        public void Translate_MovesOnlyRelativeParts(string formula, int rows, int columns, string expected) =>
            Assert.Equal(expected, References.Translate(formula, rows, columns));

        [Theory]
        // rows inserted before row index 2 (Excel row 3)
        [InlineData("A1+A3", "S", "S", true, 2, 1, "A1+A4")]
        [InlineData("SUM(A1:A5)", "S", "S", true, 2, 1, "SUM(A1:A6)")]
        [InlineData("SUM(A3:A5)", "S", "S", true, 2, 1, "SUM(A4:A6)")]
        [InlineData("SUM(A1:A2)", "S", "S", true, 2, 1, "SUM(A1:A2)")]
        [InlineData("$A$3", "S", "S", true, 2, 2, "$A$5")]
        // another sheet's formula pointing here, and here pointing elsewhere
        [InlineData("S!A3", "T", "S", true, 2, 1, "S!A4")]
        [InlineData("A3", "T", "S", true, 2, 1, "A3")]
        [InlineData("T!A3", "S", "S", true, 2, 1, "T!A3")]
        // columns inserted before column index 1 (B)
        [InlineData("A1+B1+C1", "S", "S", false, 1, 2, "A1+D1+E1")]
        [InlineData("SUM(A1:B1)", "S", "S", false, 1, 1, "SUM(A1:C1)")]
        [InlineData("SUM(A:C)", "S", "S", false, 1, 1, "SUM(A:D)")]
        [InlineData("SUM(1:3)", "S", "S", false, 1, 1, "SUM(1:3)")]
        // deletions
        [InlineData("A1+A3+A4", "S", "S", true, 2, -1, "A1+#REF!+A3")]
        [InlineData("SUM(A1:A5)", "S", "S", true, 2, -1, "SUM(A1:A4)")]
        [InlineData("SUM(A3:A5)", "S", "S", true, 2, -1, "SUM(A3:A4)")]
        [InlineData("SUM(A3:A3)", "S", "S", true, 2, -1, "SUM(#REF!)")]
        [InlineData("SUM(A1:A2)", "S", "S", true, 2, -2, "SUM(A1:A2)")]
        [InlineData("SUM(A2:A5)", "S", "S", true, 2, -3, "SUM(A2:A2)")]
        [InlineData("A1+B1", "S", "S", false, 1, -1, "A1+#REF!")]
        [InlineData("S!B1", "T", "S", false, 1, -1, "S!#REF!")]
        [InlineData("\"A3\"&A3", "S", "S", true, 2, 1, "\"A3\"&A4")]
        public void Adjust_FollowsInsertedAndDeletedRowsAndColumns(string formula, string own, string target, bool rows, int at, int delta, string expected) =>
            Assert.Equal(expected, References.Adjust(formula, own, target, rows, at, delta));

        [Theory]
        [InlineData(0, "A")] [InlineData(25, "Z")] [InlineData(26, "AA")] [InlineData(701, "ZZ")] [InlineData(702, "AAA")] [InlineData(16383, "XFD")]
        public void ColumnNames(int index, string name) => Assert.Equal(name, References.ColumnName(index));
    }
}
