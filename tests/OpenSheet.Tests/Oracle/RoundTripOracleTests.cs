using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using Xunit.Abstractions;

namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// A workbook built by Excel itself, loaded and saved again by OpenSheet, must still hold what it held. Excel counts the
    /// features of the original and of the copy; every count that changed is reported.
    /// </summary>
    [Collection("Excel")]
    public class RoundTripOracleTests : TestBase
    {
        readonly ExcelOracle _excel;
        readonly ITestOutputHelper _out;
        public RoundTripOracleTests(ExcelOracle excel, ITestOutputHelper output) { _excel = excel; _out = output; }

        void BuildRichWorkbook(string path)
        {
            File.WriteAllBytes(Tmp("pic.png"), Samples.Png);
            using (var doc = _excel.New())
            {
                dynamic wb = doc.Book;
                dynamic ws = wb.Worksheets.Item(1); ws.Name = "Data";
                ws.Range("A1").Value2 = "month"; ws.Range("B1").Value2 = "sales"; ws.Range("C1").Value2 = "region";
                var months = new[] { "Jan", "Feb", "Mar", "Apr", "May" };
                for (var i = 0; i < 5; i++) { ws.Cells(i + 2, 1).Value2 = months[i]; ws.Cells(i + 2, 2).Value2 = 10.0 * (i + 1); ws.Cells(i + 2, 3).Value2 = i % 2 == 0 ? "N" : "S"; }
                ws.Range("E1").Formula = "=SUM(B2:B6)";
                ws.Range("E2").Formula2 = "=TEXTJOIN(\",\",TRUE,A2:A6)";

                dynamic table = ws.ListObjects.Add(1, ws.Range("A1:C6"), Type.Missing, 1); table.Name = "Sales";

                dynamic co = ws.ChartObjects().Add(300, 10, 300, 200);
                co.Chart.SetSourceData(ws.Range("A1:B6"));

                ws.Range("G1:G5").FormatConditions.AddColorScale(3);
                ws.Range("H1:H5").FormatConditions.AddDatabar();
                ws.Range("I1:I5").FormatConditions.AddIconSetCondition();
                dynamic fc = ws.Range("J1:J5").FormatConditions.Add(1, 5, "=2"); fc.Interior.Color = 65535;
                for (var i = 1; i <= 5; i++) foreach (var col in new[] { "G", "H", "I", "J" }) ws.Range(col + i).Value2 = (double)i;

                ws.Range("K1:K3").Validation.Add(3, 1, 1, "x,y,z");
                wb.Names.Add("Total", "=Data!$E$1");
                ws.Range("L1").AddComment("remember");
                ws.Hyperlinks.Add(ws.Range("L2"), "https://example.com");
                ws.Range("M1:N1").Merge();
                ws.Rows("8:10").Group();
                try { ws.SparklineGroups(); } catch (Exception) { }
                try { ws.Range("P1").SparklineGroups.Add(1, "Data!B2:B6"); } catch (Exception) { }
                ws.Shapes.AddPicture(Tmp("pic.png"), false, true, 10, 250, 50, 50);
                ws.Shapes.AddShape(1, 100, 250, 60, 40);

                try
                {
                    dynamic cache = wb.PivotCaches().Create(1, "Data!A1:C6");
                    cache.CreatePivotTable(ws.Range("R1"), "Pivot");
                }
                catch (Exception) { }

                dynamic hidden = wb.Worksheets.Add(Type.Missing, ws); hidden.Name = "Hidden";
                hidden.Range("A1").Value2 = "secret"; hidden.Visible = 0;
                ws.Tab.Color = 255;
                ws.Activate();
                wb.SaveAs(path, 51);
            }
        }

        static SortedDictionary<string, string> Inventory(dynamic wb)
        {
            var inv = new SortedDictionary<string, string>();
            inv["sheets"] = string.Join(",", Enumerable.Range(1, (int)wb.Worksheets.Count).Select(i => (string)wb.Worksheets.Item(i).Name + "/" + (int)wb.Worksheets.Item(i).Visible));
            inv["names"] = string.Join(",", Enumerable.Range(1, (int)wb.Names.Count).Select(i => (string)wb.Names.Item(i).Name).OrderBy(x => x));
            foreach (var sheetName in new[] { "Data", "Hidden" })
            {
                dynamic ws = wb.Worksheets.Item(sheetName);
                string P(string k) => sheetName + "." + k;
                void Try(string key, Func<object> get) { try { inv[P(key)] = Convert.ToString(get()); } catch (Exception e) { inv[P(key)] = "n/a(" + e.GetType().Name + ")"; } }
                Try("used", () => ws.UsedRange.Address);
                Try("tables", () => ws.ListObjects.Count);
                Try("charts", () => ws.ChartObjects().Count);
                Try("shapes", () => ws.Shapes.Count);
                Try("shape list", () => string.Join(" | ", Enumerable.Range(1, (int)ws.Shapes.Count).Select(i =>
                {
                    dynamic sh = ws.Shapes.Item(i);
                    string where; try { where = (string)sh.TopLeftCell.Address; } catch (Exception) { where = "?"; }
                    string name; try { name = (string)sh.Name; } catch (Exception) { name = "?"; }
                    return (int)sh.Type + ":" + name + "@" + where;
                })));
                Try("pivots", () => ws.PivotTables().Count);
                Try("hyperlinks", () => ws.Hyperlinks.Count);
                Try("comments", () => ws.Comments.Count);
                Try("tab color", () => ws.Tab.Color);
                Try("cf count", () => ws.Cells.FormatConditions.Count);
                Try("cf types", () => string.Join(",", Enumerable.Range(1, (int)ws.Cells.FormatConditions.Count).Select(i => (int)ws.Cells.FormatConditions.Item(i).Type)));
                Try("validation K1", () => ws.Range("K1").Validation.Type);
                Try("merged M1", () => ws.Range("M1").MergeArea.Address);
                Try("outline row 9", () => ws.Rows(9).OutlineLevel);
                Try("sparklines", () => ws.Cells.SparklineGroups.Count);
                Try("E1", () => ws.Range("E1").Formula);
                Try("E2", () => ws.Range("E2").Formula2);
                Try("E2 value", () => ws.Range("E2").Value2);
                Try("B3", () => ws.Range("B3").Value2);
            }
            return inv;
        }

        [ExcelFact]
        public void LoadAndSave_KeepsWhatAnExcelBuiltWorkbookHolds()
        {
            BuildRichWorkbook(Tmp("original.xlsx"));
            var book = WorkBook.Load(Tmp("original.xlsx"));
            book.DefaultWorkSheet["T18"].Value = "touched"; // a corner of the range already in use, so the used range does not grow
            book.SaveAs(Tmp("resaved.xlsx"));
            // keep both files where they can be inspected after the test folder is gone
            File.Copy(Tmp("original.xlsx"), Path.Combine(Path.GetTempPath(), "opensheet-roundtrip-original.xlsx"), true);
            File.Copy(Tmp("resaved.xlsx"), Path.Combine(Path.GetTempPath(), "opensheet-roundtrip-resaved.xlsx"), true);

            SortedDictionary<string, string> before, after;
            using (var a = _excel.Open(Tmp("original.xlsx"))) before = Inventory(a.Book);
            using (var b = _excel.Open(Tmp("resaved.xlsx"))) after = Inventory(b.Book);

            var lost = new List<string>();
            foreach (var kv in before)
                if (!after.TryGetValue(kv.Key, out var now) || now != kv.Value)
                    lost.Add($"{kv.Key}: was [{kv.Value}] now [{now}]");
            File.WriteAllText(Path.Combine(Path.GetTempPath(), "opensheet-roundtrip-oracle.txt"),
                string.Join("\n", before.Select(kv => $"{kv.Key} = {kv.Value}")) + "\n\nCHANGED:\n" + string.Join("\n", lost));
            _out.WriteLine(string.Join("\n", lost));
            Assert.True(lost.Count == 0, lost.Count + " feature(s) changed in the round trip:" + Environment.NewLine + "  " + string.Join(Environment.NewLine + "  ", lost));
        }
    }
}
