using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>Streaming reads of hand-written packages: the corners Excel's own files rarely hit, and files that are damaged or hostile.</summary>
    public class StreamingPackageTests
    {
        const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        static void Add(ZipArchive zip, string name, string text)
        {
            using (var w = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) w.Write(text);
        }

        static string Workbook() => "<workbook xmlns=\"" + Main + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";

        static string Worksheet(string rows, string prolog = "") => prolog + "<worksheet xmlns=\"" + Main + "\"><sheetData>" + rows + "</sheetData></worksheet>";

        static byte[] Package(string rows, string sharedStrings = null, string styles = null, string workbook = null, string worksheet = null)
        {
            using (var ms = new MemoryStream())
            {
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    Add(zip, "xl/workbook.xml", workbook ?? Workbook());
                    Add(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                    Add(zip, "xl/worksheets/sheet1.xml", worksheet ?? Worksheet(rows));
                    if (sharedStrings != null) Add(zip, "xl/sharedStrings.xml", "<sst xmlns=\"" + Main + "\">" + sharedStrings + "</sst>");
                    if (styles != null) Add(zip, "xl/styles.xml", "<styleSheet xmlns=\"" + Main + "\">" + styles + "</styleSheet>");
                }
                return ms.ToArray();
            }
        }

        [Fact]
        public void SharedStrings_JoinRuns_SkipPhoneticText_KeepSpaces_DecodeEntities()
        {
            var sst = "<si><r><t>Hel</t></r><r><rPr><b/></rPr><t xml:space=\"preserve\">lo </t></r><r><t>world</t></r><rPh sb=\"0\" eb=\"1\"><t>ignored</t></rPh></si>"
                + "<si><t xml:space=\"preserve\">  padded </t></si>"
                + "<si><t>a &amp; b &lt;c&gt;</t></si>"
                + "<si><t><![CDATA[<raw>]]></t></si>"
                + "<si/>"
                + "<si><t/></si>";
            var rows = "<row r=\"1\">" + string.Concat(Enumerable.Range(0, 6).Select(i => "<c r=\"" + (char)('A' + i) + "1\" t=\"s\"><v>" + i + "</v></c>")) + "</row>";
            var cells = WorkBook.StreamRows(Package(rows, sst)).Single().Cells;
            Assert.Equal(new[] { "Hello world", "  padded ", "a & b <c>", "<raw>", "", "" }, cells.Select(c => (string)c.Value).ToArray());
        }

        [Fact]
        public void InlineStrings_AreRead_AndPhoneticTextIsLeftOut()
        {
            var rows = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>inline</t></is></c>"
                + "<c r=\"B1\" t=\"inlineStr\"><is><r><t>one </t></r><r><t>two</t></r><rPh sb=\"0\" eb=\"1\"><t>x</t></rPh></is></c>"
                + "<c r=\"C1\" t=\"inlineStr\"/></row>";
            var cells = WorkBook.StreamRows(Package(rows)).Single().Cells;
            Assert.Equal(new[] { "inline", "one two", "" }, cells.Select(c => (string)c.Value).ToArray());
        }

        [Fact]
        public void Rows_WithoutAddresses_AreNumberedInOrder_AndEmptyRowsAreSkipped()
        {
            var rows = "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"C1\" t=\"b\"><v>1</v></c></row>"
                + "<row><c><v>7</v></c><c><v>8</v></c></row>"
                + "<row r=\"5\"/>"
                + "<row><c r=\"B6\" t=\"e\"><v>#DIV/0!</v></c></row>";
            var result = WorkBook.StreamRows(Package(rows)).ToList();
            Assert.Equal(new[] { 0, 1, 5 }, result.Select(r => r.RowNumber).ToArray());
            Assert.Equal(new[] { 0, 2 }, result[0].Cells.Select(c => c.ColumnIndex).ToArray());
            Assert.Equal(StreamCellType.Boolean, result[0][2].Type);
            Assert.Equal(new[] { 0, 1 }, result[1].Cells.Select(c => c.ColumnIndex).ToArray());
            Assert.Equal(8d, result[1][1].Value);
            Assert.Equal(StreamCellType.Error, result[2][1].Type);
            Assert.Equal("#DIV/0!", result[2][1].Text);
        }

        [Fact]
        public void Cells_FirstValueWins_AndFormulasWithoutResultAreBlank()
        {
            var rows = "<row r=\"1\"><c r=\"A1\"><v>1</v><v>2</v></c><c r=\"B1\"><f>A1+1</f></c><c r=\"C1\"><f>A1+1</f><v>2</v></c><c r=\"D1\"><v/></c></row>";
            var row = WorkBook.StreamRows(Package(rows)).Single();
            Assert.Equal(1d, row[0].Value);
            Assert.Equal(StreamCellType.Blank, row[1].Type);
            Assert.True(row[1].IsFormula);
            Assert.True(row[2].IsFormula);
            Assert.Equal(2d, row[2].Value);
            Assert.Null(row[3]);
        }

        [Fact]
        public void DateStyles_TurnNumbersIntoDates()
        {
            var styles = "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\"/></numFmts>"
                + "<cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/><xf numFmtId=\"164\"/></cellXfs>";
            var rows = "<row r=\"1\"><c r=\"A1\" s=\"0\"><v>45292</v></c><c r=\"B1\" s=\"1\"><v>45292</v></c><c r=\"C1\" s=\"2\"><v>45292</v></c><c r=\"D1\" s=\"9\"><v>45292</v></c></row>";
            var row = WorkBook.StreamRows(Package(rows, styles: styles)).Single();
            Assert.Equal(StreamCellType.Number, row[0].Type);
            Assert.Equal(new DateTime(2024, 1, 1), row[1].Value);
            Assert.Equal(new DateTime(2024, 1, 1), row[2].Value);
            Assert.Equal(StreamCellType.Number, row[3].Type); // a style the file does not have is not a date
        }

        [Fact]
        public void DateStyles_OnNumbersNoDateCanHold_LeaveThemNumbers()
        {
            var styles = "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs>";
            var rows = "<row r=\"1\"><c r=\"A1\" s=\"1\"><v>-5</v></c><c r=\"B1\" s=\"1\"><v>2958466</v></c><c r=\"C1\" s=\"1\"><v>2958465.5</v></c><c r=\"D1\" s=\"1\"><v>0.5</v></c></row>";
            var row = WorkBook.StreamRows(Package(rows, styles: styles)).Single();
            Assert.Equal(StreamCellType.Number, row[0].Type);
            Assert.Equal(-5d, row[0].Value);
            Assert.Equal(StreamCellType.Number, row[1].Type);
            Assert.Equal(StreamCellType.Number, row[2].Type);
            Assert.Equal(StreamCellType.Date, row[3].Type);
        }

        [Fact]
        public void DuplicateNumberFormatIds_DoNotBreakReading()
        {
            var styles = "<numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"0\"/><numFmt numFmtId=\"164\" formatCode=\"yyyy\"/></numFmts><cellXfs count=\"1\"><xf numFmtId=\"164\"/></cellXfs>";
            var rows = "<row r=\"1\"><c r=\"A1\" s=\"0\"><v>45292</v></c></row>";
            Assert.Single(WorkBook.StreamRows(Package(rows, styles: styles)));
        }

        // ---- damaged and hostile files ----

        [Theory]
        [InlineData("5")]
        [InlineData("-1")]
        [InlineData("abc")]
        public void SharedStringIndex_OutsideTheTable_IsRejectedAsInvalidData(string index)
        {
            var rows = "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>" + index + "</v></c></row>";
            Assert.Throws<InvalidDataException>(() => WorkBook.StreamRows(Package(rows, "<si><t>only</t></si>")).ToList());
        }

        [Fact]
        public void Dtd_IsRefused_InTheSheet_AndInTheSmallParts()
        {
            const string dtd = "<!DOCTYPE x [<!ENTITY boom \"boom\">]>";
            var row = "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>&boom;</v></c></row>";
            Assert.Throws<XmlException>(() => WorkBook.GetSheetNames(Package("", workbook: dtd + Workbook())));
            Assert.Throws<XmlException>(() => WorkBook.StreamRows(Package("", worksheet: Worksheet(row, dtd))).ToList());
        }

        [Fact]
        public void OversizedSmallPart_IsRefused()
        {
            // 70 million characters of white space inside workbook.xml: inflates from a few KB, and must not be loaded whole
            byte[] bomb;
            using (var ms = new MemoryStream())
            {
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                using (var w = new StreamWriter(zip.CreateEntry("xl/workbook.xml").Open(), new UTF8Encoding(false)))
                {
                    w.Write("<workbook xmlns=\"" + Main + "\">");
                    var chunk = new string(' ', 1 << 20);
                    for (var i = 0; i < 70; i++) w.Write(chunk);
                    w.Write("</workbook>");
                }
                bomb = ms.ToArray();
            }
            Assert.Throws<XmlException>(() => WorkBook.GetSheetNames(bomb));
        }

        [Fact]
        public void NotAZip_IsReportedAsUnsupported() =>
            Assert.Throws<NotSupportedException>(() => WorkBook.StreamRows(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }).ToList());
    }
}
