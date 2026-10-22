using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>The text renderings (CSV, JSON, XML, HTML) as bytes, and the text import that reads them back.</summary>
    public class TextExportTests : TestBase
    {
        static readonly UTF8Encoding Strict = new UTF8Encoding(false, true);

        static WorkSheet Sheet(params object[][] rows)
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            for (var r = 0; r < rows.Length; r++)
                for (var c = 0; c < rows[r].Length; c++)
                    if (rows[r][c] != null) ws.GetCellAt(r, c).Value = rows[r][c];
            return ws;
        }

        static string Csv(WorkSheet ws, string delimiter = ",") => Strict.GetString(ws.ToCsvStream(delimiter).ToArray());

        // ---- CSV ----

        [Fact]
        public void Csv_QuotesOnlyWhatNeedsIt()
        {
            var ws = Sheet(new object[] { "plain", "a,b", "say \"hi\"", "two\nlines", "cr\rhere", 3.5 });
            Assert.Equal("plain,\"a,b\",\"say \"\"hi\"\"\",\"two\nlines\",\"cr\rhere\",3.5\r\n", Csv(ws));
        }

        [Fact]
        public void Csv_DoublesEveryQuote_AndKeepsSurrogatePairsWhole()
        {
            var ws = Sheet(new object[] { "\"a\"\"b\"", "\U0001F600\"\U0001F600\"" });
            var text = Csv(ws); // the strict decoder throws on a split surrogate pair
            Assert.Equal("\"\"\"a\"\"\"\"b\"\"\",\"\U0001F600\"\"\U0001F600\"\"\"\r\n", text);
        }

        [Fact]
        public void Csv_GapsAreEmptyFields_AndTheUsedRangeStartsAtTheFirstUsedCell()
        {
            var ws = Sheet(new object[] { null, null, null }, new object[] { null, "x", null, "y" });
            Assert.Equal("x,,y\r\n", Csv(ws));
        }

        [Fact]
        public void Csv_QuotesByTheDelimiterInUse()
        {
            var tab = Sheet(new object[] { "a|b", "tab\there", "a,b", "x" });
            Assert.Equal("a|b\t\"tab\there\"\ta,b\tx\r\n", Strict.GetString(tab.ToTsvStream().ToArray()));
            Assert.Equal("\"a;b\";c\r\n", Csv(Sheet(new object[] { "a;b", "c" }), ";"));
        }

        [Fact]
        public void Csv_MultiCharDelimiter_QuotesFieldsThatContainIt()
        {
            Assert.Equal("\"a||b\"||a|b||c\r\n", Csv(Sheet(new object[] { "a||b", "a|b", "c" }), "||"));
        }

        [Fact]
        public void Csv_UsesDisplayText()
        {
            var ws = Sheet(new object[] { 1234.5, new DateTime(2024, 3, 9), true });
            ws["A1"].FormatString = "#,##0.00";
            ws["B1"].FormatString = "yyyy-mm-dd";
            Assert.Equal("\"1,234.50\",2024-03-09,TRUE\r\n", Csv(ws));
        }

        [Fact]
        public void Csv_OfAnEmptySheet_IsEmpty()
        {
            Assert.Empty(WorkBook.Create().DefaultWorkSheet.ToCsvStream().ToArray());
        }

        [Fact]
        public void Csv_HasNoByteOrderMark()
        {
            var bytes = Sheet(new object[] { "é" }).ToCsvStream().ToArray();
            Assert.Equal(new byte[] { 0xC3, 0xA9, 0x0D, 0x0A }, bytes);
        }

        [Fact]
        public void Csv_SaveAs_WritesTheSameBytesAsTheStream()
        {
            var ws = Sheet(new object[] { "x", "y,z" }, new object[] { 1, 2 });
            var path = Tmp("out.csv");
            ws.Book.SaveAs(path);
            Assert.Equal(ws.ToCsvStream().ToArray(), File.ReadAllBytes(path));
        }

        [Fact]
        public void Csv_ReadsBack_ThroughTheImporter()
        {
            var path = Tmp("in.csv");
            File.WriteAllText(path, "id;name;note\r\n1;\"Smith; John\";\"line1\nline2\"\r\n2;\"say \"\"hi\"\"\";\r\n");
            var ws = WorkBook.LoadCSV(path).DefaultWorkSheet;
            Assert.Equal("Smith; John", ws["B2"].StringValue);
            Assert.Equal("line1\nline2", ws["C2"].StringValue);
            Assert.Equal("say \"hi\"", ws["B3"].StringValue);
            Assert.Equal(2d, ws["A3"].DoubleValue);
            Assert.True(ws["C3"].IsEmpty);
        }

        // ---- CSV import ----

        [Fact]
        public void Import_Csv_GuessesDelimiter_SkipsBlankLines_KeepsLeadingZeros()
        {
            var path = Tmp("guess.csv");
            File.WriteAllText(path, "\r\n\r\n a;b;c\r\n\r\n00123;4,5;true\r\n");
            var ws = WorkBook.LoadCSV(path).DefaultWorkSheet;
            Assert.Equal(" a", ws["A1"].StringValue);
            Assert.Equal("00123", ws["A2"].StringValue);
            Assert.True(ws.GetCellAt(1, 0).IsText);
            Assert.Equal("4,5", ws["B2"].StringValue);
            Assert.True(ws.GetCellAt(1, 2).IsBoolean);
            Assert.Equal(2, ws.PhysicalRowCount);
        }

        [Fact]
        public void Import_Csv_MultiCharDelimiter_AndWhitespaceOnlyCells()
        {
            var path = Tmp("multi.csv");
            File.WriteAllText(path, "a<>b<> <>d\r\n1<>2<><>4");
            var ws = WorkBook.LoadCSV(path, ExcelFileFormat.XLSX, "<>").DefaultWorkSheet;
            Assert.Equal("b", ws["B1"].StringValue);
            Assert.True(ws["C1"].IsEmpty);
            Assert.Equal("d", ws["D1"].StringValue);
            Assert.Equal(4d, ws["D2"].DoubleValue);
        }

        [Fact]
        public void Import_Csv_ConvertFieldValuesOff_KeepsEverythingText()
        {
            var path = Tmp("text.csv");
            File.WriteAllText(path, "1,true,2024-01-02\r\n");
            var ws = WorkBook.LoadCSV(path, ExcelFileFormat.XLSX, null, false).DefaultWorkSheet;
            Assert.True(ws.GetCellAt(0, 0).IsText && ws.GetCellAt(0, 1).IsText && ws.GetCellAt(0, 2).IsText);
        }

        // ---- JSON ----

        [Fact]
        public void Json_IsAnArrayOfRows_WithNullForGaps()
        {
            var ws = Sheet(new object[] { "a", 1.5, true }, new object[] { null, new DateTime(2024, 1, 2, 3, 4, 5), "é<>&" });
            using (var doc = JsonDocument.Parse(ws.ToJsonStream().ToArray()))
            {
                var rows = doc.RootElement.EnumerateArray().Select(r => r.EnumerateArray().ToArray()).ToArray();
                Assert.Equal("a", rows[0][0].GetString());
                Assert.Equal(1.5, rows[0][1].GetDouble());
                Assert.True(rows[0][2].GetBoolean());
                Assert.Equal(JsonValueKind.Null, rows[1][0].ValueKind);
                Assert.Equal("2024-01-02T03:04:05", rows[1][1].GetString());
                Assert.Equal("é<>&", rows[1][2].GetString());
            }
        }

        [Fact]
        public void Json_WorkbookIsAnObjectOfSheets()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = 1;
            wb.CreateWorkSheet("Other")["A1"].Value = "x";
            var path = Tmp("book.json");
            wb.SaveAs(path);
            using (var doc = JsonDocument.Parse(File.ReadAllBytes(path)))
                Assert.Equal(new[] { "Sheet1", "Other" }, doc.RootElement.EnumerateObject().Select(p => p.Name).ToArray());
        }

        // ---- XML ----

        [Fact]
        public void Xml_BytesParse_AndDeclareTheirOwnEncoding()
        {
            var ws = Sheet(new object[] { "café", 2.5, true }, new object[] { null, null, new DateTime(2024, 1, 2) });
            var bytes = ws.ToXmlStream().ToArray();
            var doc = XDocument.Load(new MemoryStream(bytes)); // fails when the declaration says utf-16 over UTF-8 bytes
            Assert.Equal("utf-8", doc.Declaration.Encoding, true);
            var cells = doc.Descendants("cell").ToList();
            Assert.Equal(new[] { "A1", "B1", "C1", "C2" }, cells.Select(c => (string)c.Attribute("ref")).ToArray());
            Assert.Equal(new[] { "string", "number", "boolean", "date" }, cells.Select(c => (string)c.Attribute("type")).ToArray());
            Assert.Equal("café", cells[0].Value);
            Assert.Equal("2.5", cells[1].Value);
            Assert.Equal((byte)'<', bytes[0]); // no byte-order mark
        }

        [Fact]
        public void Xml_ReadsBack_ThroughTheImporter()
        {
            var ws = Sheet(new object[] { "x", 2.5, true });
            var path = Tmp("rt.xml");
            ws.Book.SaveAs(path);
            var back = WorkBook.Load(path).DefaultWorkSheet;
            Assert.Equal("x", back["A1"].StringValue);
            Assert.Equal(2.5, back["B1"].DoubleValue);
            Assert.True(back["C1"].BoolValue);
        }

        [Fact]
        public void Import_Xml_RefusesADocType_SoEntityBombsNeverExpand()
        {
            var path = Tmp("bomb.xml");
            File.WriteAllText(path, "<?xml version=\"1.0\"?><!DOCTYPE workbook [<!ENTITY a \"aaaaaaaaaa\"><!ENTITY b \"&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;\">]>"
                + "<workbook><sheet name=\"S\"><row index=\"1\"><cell ref=\"A1\" type=\"string\">&b;</cell></row></sheet></workbook>");
            Assert.Throws<System.Xml.XmlException>(() => WorkBook.Load(path));
        }

        // ---- HTML ----

        [Fact]
        public void Html_EncodesText_AndKeepsLeadingSpaces()
        {
            var ws = Sheet(new object[] { "   lead", "<b>&\"'", "tail  " });
            var html = ws.Book.ExportToHtmlString();
            Assert.Contains("<td>&nbsp;&nbsp;&nbsp;lead</td>", html);
            Assert.Contains("<td>&lt;b&gt;&amp;&quot;&#39;</td>", html);
            Assert.Contains("<td>tail  </td>", html);
            Assert.Contains("<th>A</th><th>B</th><th>C</th>", html);
            Assert.Contains("<th>1</th>", html);
        }

        [Fact]
        public void Html_Options_ControlSpacesHeadersAndHiddenColumns()
        {
            var ws = Sheet(new object[] { "  a", "b", "c" });
            ws.Native.SetColumnHidden(1, true);
            var html = ws.Book.ExportToHtmlString(outputColumnHeaders: false, outputHiddenColumns: false, outputLeadingSpacesAsNonBreaking: false, outputRowNumbers: false);
            Assert.Contains("<tr><td>  a</td><td>c</td></tr>", html);
            Assert.DoesNotContain("<th>", html);
            Assert.Contains("<td>b</td>", ws.Book.ExportToHtmlString(outputHiddenColumns: true));
        }

        [Fact]
        public void Html_EmptySheet_HasNoRows()
        {
            var html = WorkBook.Create().ExportToHtmlString(outputColumnHeaders: false, outputRowNumbers: false);
            Assert.DoesNotContain("<tr>", html);
        }
    }
}
