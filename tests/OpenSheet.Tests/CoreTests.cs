using System;
using System.IO;
using System.Linq;
using OpenSheet.Styles;
using Xunit;

namespace OpenSheet.Tests
{
    public class CoreTests : TestBase
    {
        [Fact]
        public void HelloWorld_RoundTripsThroughXlsx()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = "Hello World";
            wb.SaveAs(Tmp("a.xlsx"));

            var back = WorkBook.Load(Tmp("a.xlsx"));
            Assert.Equal("Hello World", back.DefaultWorkSheet["A1"].StringValue);
            Assert.Equal(Path.GetFullPath(Tmp("a.xlsx")), back.FilePath);
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX, "x.xlsx")]
        [InlineData(ExcelFileFormat.XLS, "x.xls")]
        public void ValueTypes_RoundTrip(ExcelFileFormat fmt, string file)
        {
            var wb = WorkBook.Create(fmt);
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "text";
            ws["A2"].Value = 42;
            ws["A3"].Value = 3.5;
            ws["A4"].Value = true;
            ws["A5"].Value = new DateTime(2024, 5, 17, 13, 30, 0);
            ws["A6"].Value = 12.34m;
            wb.SaveAs(Tmp(file));

            var s = WorkBook.Load(Tmp(file)).DefaultWorkSheet;
            Assert.Equal("text", s["A1"].Value);
            Assert.Equal(42d, s["A2"].Value);
            Assert.Equal(3.5, s["A3"].Value);
            Assert.Equal(true, s["A4"].Value);
            Assert.Equal(new DateTime(2024, 5, 17, 13, 30, 0), s["A5"].Value);
            Assert.True(s.GetCellAt(4, 0).IsDateTime);
            Assert.Equal(12.34m, s["A6"].DecimalValue);
            Assert.True(s.GetCellAt(0, 0).IsText);
            Assert.True(s.GetCellAt(1, 0).IsNumeric);
            Assert.True(s.GetCellAt(3, 0).IsBoolean);
        }

        [Fact]
        public void NullAndEmptyAndWhitespace_RemoveTheCell()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = "x"; ws["A2"].Value = "y"; ws["A3"].Value = "z";
            ws["A1"].Value = null; ws["A2"].Value = ""; ws["A3"].Value = "   ";
            Assert.True(ws.GetCellAt(0, 0).IsEmpty);
            Assert.True(ws.GetCellAt(1, 0).IsEmpty);
            Assert.True(ws.GetCellAt(2, 0).IsEmpty);
            Assert.Equal(0, ws.PhysicalColumnCount);
        }

        [Fact]
        public void Formulas_AreEvaluated_AndStayFormulas()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = 2; ws["A2"].Value = 3;
            ws["A3"].Formula = "=A1*A2+1";
            ws["B1"].Formula = "SUM(A1:A2)";
            Assert.Equal(7d, ws["A3"].Value);
            Assert.Equal("=A1*A2+1", ws["A3"].Formula);
            Assert.Equal(5d, ws["B1"].DoubleValue);
            Assert.True(ws.GetCellAt(2, 0).IsFormula);

            ws["A1"].Value = 10; // edit invalidates cached results
            Assert.Equal(31d, ws["A3"].Value);

            ws.GetCellAt(2, 0).RemoveFormula();
            Assert.False(ws.GetCellAt(2, 0).IsFormula);
        }

        [Fact]
        public void FormulaErrors_AreReportedAsErrors()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Formula = "=1/0";
            var c = ws.GetCellAt(0, 0);
            Assert.True(c.IsError);
            Assert.Equal("#DIV/0!", c.StringValue);
        }

        [Fact]
        public void Range_Aggregates()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = 1; ws["A2"].Value = 2; ws["A3"].Value = 6; ws["A4"].Value = "skip";
            var r = ws["A1:A4"];
            Assert.Equal(9m, r.Sum());
            Assert.Equal(6m, r.Max());
            Assert.Equal(1m, r.Min());
            Assert.Equal(3m, r.Avg());
            Assert.Equal(9m, r.Value);
            Assert.Equal(4, r.Count);
            Assert.Equal(1, r.ColumnCount);
            Assert.Equal(4, r.RowCount);
        }

        [Fact]
        public void Range_SetValue_WritesEveryCell()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["B2:C3"].Value = 7;
            Assert.Equal(28m, ws["B2:C3"].Sum());
            ws["B2:C3"].Value = null;
            Assert.True(ws["B2:C3"].IsEmpty);
        }

        [Fact]
        public void Address_Parsing_And_Errors()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            Assert.Equal("B2:D9", ws["D9:B2"].RangeAddressAsString);
            Assert.Equal("C3", ws["$C$3"].RangeAddressAsString);
            Assert.Throws<ArgumentException>(() => ws["Z"]);
            Assert.Throws<ArgumentException>(() => ws["A0"]);
            Assert.Throws<ArgumentException>(() => ws["A1:B2:C3"]);
        }

        [Fact]
        public void Styles_Persist_InXlsx()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "styled";
            ws["A1"].Style.Font.Bold = true;
            ws["A1"].Style.Font.Height = 16;
            ws["A1"].Style.Font.SetColor("#FF0000");
            ws["A1"].Style.SetBackgroundColor("#00FF00");
            ws["A1"].Style.BottomBorder.SetBorderStyle("#0000FF", BorderType.Thick);
            ws["A1"].Style.HorizontalAlignment = HorizontalAlignment.Center;
            ws["A1"].Style.WrapText = true;
            ws["A1"].FormatString = "0.00%";
            wb.SaveAs(Tmp("s.xlsx"));

            var st = WorkBook.Load(Tmp("s.xlsx")).DefaultWorkSheet["A1"].Style;
            Assert.True(st.Font.Bold);
            Assert.Equal(16, st.Font.Height);
            Assert.Equal("#FF0000", st.Font.Color);
            Assert.Equal("#00FF00", st.BackgroundColor);
            Assert.Equal(FillPattern.SolidForeground, st.FillPattern);
            Assert.Equal(BorderType.Thick, st.BottomBorder.Type);
            Assert.Equal("#0000FF", st.BottomBorder.Color);
            Assert.Equal(HorizontalAlignment.Center, st.HorizontalAlignment);
            Assert.True(st.WrapText);
            Assert.Equal("0.00%", WorkBook.Load(Tmp("s.xlsx")).DefaultWorkSheet["A1"].FormatString);
        }

        [Fact]
        public void Styles_BulkEdit_DoesNotMultiplyStyles()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            var before = wb.Native.NumCellStyles;
            ws["A1:J100"].Style.Font.Bold = true;
            Assert.True(wb.Native.NumCellStyles - before <= 1, "one style for 1000 identical edits");
            Assert.True(ws["J100"].Style.Font.Bold);
        }

        [Fact]
        public void Styles_Xls_UsesNearestPaletteColor()
        {
            var wb = WorkBook.Create(ExcelFileFormat.XLS);
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = 1;
            ws["A1"].Style.SetBackgroundColor("#FF0000");
            ws["A1"].Style.Font.Bold = true;
            wb.SaveAs(Tmp("s.xls"));
            var st = WorkBook.Load(Tmp("s.xls")).DefaultWorkSheet["A1"].Style;
            Assert.True(st.Font.Bold);
            Assert.Equal("#FF0000", st.BackgroundColor);
        }

        [Fact]
        public void Csv_RoundTrip_WithQuotingAndTypes()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "name"; ws["B1"].Value = "note"; ws["C1"].Value = "n";
            ws["A2"].Value = "Ann, Jr."; ws["B2"].Value = "says \"hi\"\nthere"; ws["C2"].Value = 12.5;
            wb.SaveAs(Tmp("d.csv"));

            var text = File.ReadAllText(Tmp("d.csv"));
            Assert.Contains("\"Ann, Jr.\"", text);
            Assert.Contains("\"says \"\"hi\"\"\nthere\"", text);

            var back = WorkBook.Load(Tmp("d.csv")).DefaultWorkSheet;
            Assert.Equal("Ann, Jr.", back["A2"].Value);
            Assert.Equal("says \"hi\"\nthere", back["B2"].Value);
            Assert.Equal(12.5, back["C2"].Value);
        }

        [Fact]
        public void Csv_Import_KeepsLeadingZeros_AndDetectsSemicolons()
        {
            File.WriteAllText(Tmp("z.csv"), "zip;qty;flag;when\n00123;5;true;2024-01-31\n");
            var s = WorkBook.Load(Tmp("z.csv")).DefaultWorkSheet;
            Assert.Equal("00123", s["A2"].Value);
            Assert.Equal(5d, s["B2"].Value);
            Assert.Equal(true, s["C2"].Value);
            Assert.Equal(new DateTime(2024, 1, 31), s["D2"].Value);
        }

        [Fact]
        public void Encryption_RoundTrip_AndWrongPassword()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = "secret";
            wb.SaveAs(Tmp("e.xlsx"), "pw123");

            Assert.Equal("secret", WorkBook.Load(Tmp("e.xlsx"), "pw123").DefaultWorkSheet["A1"].Value);
            Assert.Throws<InvalidDataException>(() => WorkBook.Load(Tmp("e.xlsx"), "nope"));
            Assert.False(WorkBook.TryLoad(Tmp("e.xlsx"), "nope", out _));
            Assert.True(WorkBook.TryLoad(Tmp("e.xlsx"), "pw123", out var ok));
            Assert.True(ok.IsEncrypted); // re-saving must not silently drop protection
        }

        [Fact]
        public void Encryption_UsesAes256Sha512_AndIsReallyEncrypted()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = "plain-text-marker";
            wb.SaveAs(Tmp("k.xlsx"), "pw");
            var bytes = File.ReadAllBytes(Tmp("k.xlsx"));
            Assert.Equal(0xD0, bytes[0]);                                   // OLE2 container, not a zip
            Assert.DoesNotContain("plain-text-marker", System.Text.Encoding.ASCII.GetString(bytes));
            Assert.DoesNotContain("plain-text-marker", System.Text.Encoding.Unicode.GetString(bytes));

            // the EncryptionInfo stream is UTF-8 XML after an 8-byte header: look for the chosen algorithm
            var text = System.Text.Encoding.UTF8.GetString(bytes);
            Assert.Contains("keyBits=\"256\"", text);
            Assert.Contains("hashAlgorithm=\"SHA512\"", text);
        }

        [Fact]
        public void Xls_Password_IsRejected_NotIgnored()
        {
            var wb = WorkBook.Create(ExcelFileFormat.XLS);
            Assert.Throws<NotSupportedException>(() => wb.SaveAs(Tmp("p.xls"), "pw"));
        }

        [Fact]
        public void SaveAs_ConvertsBetweenFormats()
        {
            var wb = WorkBook.Create();
            wb.DefaultWorkSheet["A1"].Value = "x"; wb.DefaultWorkSheet["B1"].Value = 5;
            wb.SaveAs(Tmp("c.xls"));
            var back = WorkBook.Load(Tmp("c.xls"));
            Assert.Equal("x", back.DefaultWorkSheet["A1"].Value);
            Assert.Equal(5d, back.DefaultWorkSheet["B1"].Value);
        }

        [Fact]
        public void UnknownExtension_Throws()
        {
            Assert.Throws<NotSupportedException>(() => WorkBook.Create().SaveAs(Tmp("x.dat")));
        }

        [Fact]
        public void Load_MissingFile_Throws_And_TryLoad_ReturnsFalse()
        {
            Assert.Throws<FileNotFoundException>(() => WorkBook.Load(Tmp("nope.xlsx")));
            Assert.False(WorkBook.TryLoad(Tmp("nope.xlsx"), out var wb));
            Assert.Null(wb);
        }

        [Fact]
        public void Sheets_CreateFindRemoveReorder()
        {
            var wb = WorkBook.Create();
            wb.CreateWorkSheet("Two"); wb.CreateWorkSheet("Three");
            Assert.Equal(3, wb.WorkSheets.Count);
            Assert.Same(wb.GetWorkSheet("two"), wb.WorkSheets[1]); // case-insensitive, same instance
            Assert.Null(wb.GetWorkSheet("missing"));
            wb.SetSheetPosition("Three", 0);
            Assert.Equal("Three", wb.WorkSheets[0].Name);
            wb.RemoveWorkSheet("Two");
            Assert.Equal(new[] { "Three", "Sheet1" }, wb.WorkSheets.Select(s => s.Name).ToArray());
            wb.WorkSheets[0].Name = "Renamed";
            Assert.NotNull(wb.GetWorkSheet("Renamed"));
        }

        [Fact]
        public void InsertAndRemoveRows()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = "r1"; ws["A2"].Value = "r2"; ws["A3"].Value = "r3";
            ws.InsertRow(1);
            Assert.Equal("r1", ws["A1"].Value);
            Assert.True(ws.GetCellAt(1, 0).IsEmpty);
            Assert.Equal("r2", ws["A3"].Value);
            Assert.Equal("r3", ws["A4"].Value);

            ws.RemoveRow(1); // empty row -> rows below move up
            Assert.Equal("r2", ws["A2"].Value);
            Assert.Equal("r3", ws["A3"].Value);
        }

        [Fact]
        public void InsertAndRemoveColumns()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = "a"; ws["B1"].Value = "b"; ws["C1"].Value = "c";
            ws.InsertColumn(1);
            Assert.Equal("a", ws["A1"].Value);
            Assert.True(ws.GetCellAt(0, 1).IsEmpty);
            Assert.Equal("b", ws["C1"].Value);
            Assert.Equal("c", ws["D1"].Value);
            ws.RemoveColumn(1);
            Assert.Equal("b", ws["B1"].Value);
            Assert.Equal("c", ws["C1"].Value);
        }

        [Fact]
        public void SortByColumn_KeepsRowsTogether()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var data = new object[][] { new object[] { "b", 2 }, new object[] { "c", 3 }, new object[] { "a", 1 }, new object[] { null, 9 } };
            for (var r = 0; r < data.Length; r++) { ws.GetCellAt(r, 0).Value = data[r][0]; ws.GetCellAt(r, 1).Value = data[r][1]; }
            ws["A1:B4"].SortByColumn(0, SortOrder.Ascending);
            Assert.Equal(new object[] { "a", "b", "c", null }, Enumerable.Range(0, 4).Select(r => ws.GetCellAt(r, 0).Value).ToArray());
            Assert.Equal(new object[] { 1d, 2d, 3d, 9d }, Enumerable.Range(0, 4).Select(r => ws.GetCellAt(r, 1).Value).ToArray());
            ws["A1:B4"].SortByColumn("B", SortOrder.Descending);
            Assert.Equal(9d, ws["B1"].Value);
            Assert.Equal(1d, ws["B4"].Value);
        }

        [Fact]
        public void MergeAndFreeze_Persist()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "wide";
            ws.Merge("A1:C1");
            ws.CreateFreezePane(1, 1);
            wb.SaveAs(Tmp("m.xlsx"));
            var back = WorkBook.Load(Tmp("m.xlsx")).DefaultWorkSheet;
            Assert.Equal("A1:C1", back.GetMergedRegions().Single().RangeAddressAsString);
            back.Unmerge("A1:C1");
            Assert.Empty(back.GetMergedRegions());
        }

        [Fact]
        public void NamedRanges()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1:A3"].Value = 2;
            ws["A1:A3"].SaveAsNamedRange("Twos", true);
            ws.AddNamedRange("Local", ws["B1"], false);
            Assert.Equal("'Sheet1'!$A$1:$A$3", ws.FindNamedRange("Twos"));
            Assert.Contains("Twos", ws.GetNamedRanges());
            Assert.Contains("Local", ws.GetNamedRanges());
            ws["C1"].Formula = "=SUM(Twos)";
            Assert.Equal(6d, ws["C1"].Value);
            ws.RemoveNamedRange("Twos");
            Assert.DoesNotContain("Twos", ws.GetNamedRanges());
        }

        [Fact]
        public void Hyperlink_Comment_AndHeaderFooter()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "link";
            ws.GetCellAt(0, 0).Hyperlink = "https://example.com";
            ws.GetCellAt(0, 0).AddComment("note", "me", true);
            ws.Header.Center = "Title";
            ws.Footer.Right = "Page";
            ws.PrintSetup.PrintOrientation = Printing.PrintOrientation.Landscape;
            ws.PrintSetup.PaperSize = Printing.PaperSize.A4;
            wb.SaveAs(Tmp("h.xlsx"));
            var b = WorkBook.Load(Tmp("h.xlsx")).DefaultWorkSheet;
            Assert.Equal("https://example.com", b.GetCellAt(0, 0).Hyperlink);
            Assert.Equal("note", b.GetCellAt(0, 0).Comment.Content);
            Assert.Equal("me", b.GetCellAt(0, 0).Comment.Author);
            Assert.Equal("Title", b.Header.Center);
            Assert.Equal("Page", b.Footer.Right);
            Assert.Equal(Printing.PrintOrientation.Landscape, b.PrintSetup.PrintOrientation);
            Assert.Equal(Printing.PaperSize.A4, b.PrintSetup.PaperSize);
        }

        [Fact]
        public void Copy_RangeAndSheet()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = 1; ws["B1"].Value = "two"; ws["A1"].Style.Font.Bold = true;
            var other = wb.CreateWorkSheet("Other");
            ws["A1:B1"].Copy(other, "C3");
            Assert.Equal(1d, other["C3"].Value);
            Assert.Equal("two", other["D3"].Value);
            Assert.True(other["C3"].Style.Font.Bold);

            var dup = ws.CopySheet("Dup");
            Assert.Equal("two", dup["B1"].Value);
            Assert.Equal(3, wb.WorkSheets.Count);
        }

        [Fact]
        public void ReplaceText_AndNumbers()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = "foo bar foo"; ws["A2"].Value = 150;
            ws["A1:A2"].Replace("foo", "baz");
            ws["A1:A2"].Replace("15", "27");
            Assert.Equal("baz bar baz", ws["A1"].Value);
            Assert.Equal(270d, ws["A2"].Value);
        }

        [Fact]
        public void DataTable_RoundTrip()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = "id"; ws["B1"].Value = "name";
            ws["A2"].Value = 1; ws["B2"].Value = "x";
            ws["A3"].Value = 2; ws["B3"].Value = null;
            var t = ws["A1:B3"].ToDataTable(true);
            Assert.Equal(new[] { "id", "name" }, t.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToArray());
            Assert.Equal(typeof(double), t.Columns[0].DataType);
            Assert.Equal(typeof(string), t.Columns[1].DataType);
            Assert.True(t.Rows[1].IsNull(1));

            var wb2 = WorkBook.Load(new System.Data.DataSet { Tables = { t } });
            Assert.Equal("id", wb2.DefaultWorkSheet["A1"].Value);
            Assert.Equal(2d, wb2.DefaultWorkSheet["A3"].Value);
        }

        [Fact]
        public void Json_Xml_Html_Exports_And_Reimport()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws["A1"].Value = "k"; ws["B1"].Value = 1.5; ws["C1"].Value = true;
            wb.SaveAsJson(Tmp("o.json")); wb.SaveAsXml(Tmp("o.xml"));
            wb.ExportToHtml(Tmp("o.html"));
            Assert.Contains("\"k\"", File.ReadAllText(Tmp("o.json")));
            Assert.Contains("<td>k</td>", File.ReadAllText(Tmp("o.html")));
            foreach (var f in new[] { "o.json", "o.xml" })
            {
                var back = WorkBook.Import(Tmp(f)).DefaultWorkSheet;
                Assert.Equal("k", back["A1"].Value);
                Assert.Equal(1.5, back["B1"].Value);
                Assert.Equal(true, back["C1"].Value);
            }
        }

        [Fact]
        public void RangeAddress_Operations()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var a = ws["A1:B2"].RangeAddress; var b = ws["B2:C3"].RangeAddress;
            Assert.Equal("B2", a.GetIntersection(b).Location);
            Assert.Equal("A1:C3", a.Extend(b).Location);
            Assert.Null(a.GetIntersection(ws["D4:E5"].RangeAddress));
            Assert.True(a.Includes(1, 1));
            Assert.False(a.Includes(2, 2));
            var r = ws["A1:B2"]; r.AddRange(ws["C3:D4"]);
            Assert.Equal("A1:D4", r.RangeAddressAsString);
            var s = ws["A1:B2"]; s.SubtractRange(ws["A1:B1"]);
            Assert.Equal("A2:B2", s.RangeAddressAsString);
        }

        [Fact]
        public void Typed_Getters_And_GetValue()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            ws["A1"].Value = 2.9; ws["A2"].Value = "text"; ws["A3"].Value = 5;
            Assert.Equal(2, ws.GetCellAt(0, 0).IntValue);          // truncates
            Assert.Equal(0, ws.GetCellAt(1, 0).IntValue);          // non-numeric -> 0
            Assert.Equal(5, ws.GetCellAt(2, 0).GetValue<int>());
            Assert.Equal("5", ws.GetCellAt(2, 0).GetValue<string>());
            Assert.False(ws.GetCellAt(1, 0).TryGetValue<int>(out _));
            Assert.Throws<FormatException>(() => ws.GetCellAt(1, 0).GetValue<int>());
        }
    }
}
