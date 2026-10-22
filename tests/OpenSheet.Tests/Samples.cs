using System;
using OpenSheet.DataValidations;
using OpenSheet.Drawing.Charts;
using OpenSheet.Drawing.Images;
using OpenSheet.Styles;
using CfOp = OpenSheet.Formatting.Enums.ComparisonOperator;
using DvOp = OpenSheet.DataValidations.ComparisonOperator;

namespace OpenSheet.Tests
{
    /// <summary>Workbooks that use as many features as possible; shared by the schema and Excel checks.</summary>
    public static class Samples
    {
        public static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

        /// <summary>Sheet "Data" with every kind of feature, and sheet "Charts" with six charts. Positions are asserted by the tests.</summary>
        public static WorkBook KitchenSink()
        {
            var wb = WorkBook.Create();
            var ws = wb.DefaultWorkSheet;
            ws.Name = "Data";
            ws["A1"].Value = "month"; ws["B1"].Value = "sales"; ws["C1"].Value = "cost";
            var months = new[] { "Jan", "Feb", "Mar", "Apr", "May" };
            for (var i = 0; i < 5; i++) { ws.GetCellAt(i + 1, 0).Value = months[i]; ws.GetCellAt(i + 1, 1).Value = 10 * (i + 1); ws.GetCellAt(i + 1, 2).Value = 6 * (i + 1); }

            ws["A1:C1"].Style.Font.Bold = true;
            ws.Merge("E1:G1"); ws["E1"].Value = "merged";
            ws.CreateFreezePane(1, 1);
            ws.SetAutoFilter("O1:Q6"); // not A1:C6: a sheet filter and a table cannot share cells (see the table below)
            ws.SetRowBreak(10); ws.SetColumnBreak(5);
            ws.GroupRows(2, 3);
            ws.SetPrintArea("A1:G20"); ws.SetRepeatingRows(0, 0);
            ws.PrintSetup.PrintOrientation = Printing.PrintOrientation.Landscape; ws.PrintSetup.PaperSize = Printing.PaperSize.A4;
            ws.Header.Center = "Report"; ws.Footer.Right = "Page";
            ws.SetZoom(120);
            ws.SetWorksheetTabColor(new byte[] { 255, 0, 0 });
            ws["A1:A6"].SaveAsNamedRange("Months", true);
            ws.GetCellAt(0, 0).Hyperlink = "https://example.com";
            ws.GetCellAt(1, 1).AddComment("check this", "me", false);
            ws.GetColumn(0).Width = 4000;

            ws.DataValidations.AddStringListRule("H1:H5", new[] { "a", "b" });
            ws.DataValidations.AddIntegerRule("I1:I5", DvOp.Between, 1, 10).ErrorBoxText = "1-10";
            ws.DataValidations.AddDecimalRule("J1", DvOp.GreaterThan, 0.5m, null);
            ws.DataValidations.AddDateRule("K1", DvOp.GreaterThanOrEqual, new DateTime(2024, 1, 1), null);
            ws.DataValidations.AddTextLengthRule("L1", DvOp.LessThanOrEqual, 8, null);
            ws.DataValidations.AddFormulaListRule("M1", "$A$2:$A$6");

            var cf = ws.ConditionalFormatting;
            var rule = cf.CreateConditionalFormattingRule(CfOp.GreaterThan, "20");
            rule.FontFormatting.IsBold = true; rule.FontFormatting.FontColor = "#FF0000";
            rule.PatternFormatting.BackgroundColor = "#FFFF00";
            cf.AddConditionalFormatting("B2:B6", rule);

            ws.InsertImage(Png, ImageFormat.PNG, 9, 8, 11, 10);
            // The table sits away from the cells the charts plot: Excel drops literal series names from charts that read table columns.
            ws["R1"].Value = "k"; ws["S1"].Value = "v"; ws["T1"].Value = "w";
            for (var i = 2; i <= 4; i++) { ws.GetCellAt(i - 1, 17).Value = "row" + i; ws.GetCellAt(i - 1, 18).Value = i; ws.GetCellAt(i - 1, 19).Value = i * 1.5; }
            ws.AddNamedTable("Sales", ws["R1:T4"], true, TableStyle.TableStyleMedium2, true);

            wb.Metadata.Author = "tester"; wb.Metadata.Title = "Sink";
            wb.Metadata.CustomProperties.AddProperty("k", "v");

            var chartsSheet = wb.CreateWorkSheet("Charts");
            var row = 0;
            foreach (var type in new[] { ChartType.Column, ChartType.Bar, ChartType.Line, ChartType.Area, ChartType.Pie })
            {
                var c = chartsSheet.CreateChart(type, row, 0, row + 14, 8);
                c.AddSeries("Data!A2:A6", "Data!B2:B6").Title = "Sales";
                if (type != ChartType.Pie) c.AddSeries("Data!A2:A6", "Data!C2:C6").Title = "Cost";
                c.SetTitle(type.ToString()); c.SetLegendPosition(LegendPosition.Bottom); c.Plot();
                row += 16;
            }
            var sc = chartsSheet.CreateChart(ChartType.Scatter, row, 0, row + 14, 8);
            sc.AddSeries("Data!B2:B6", "Data!C2:C6"); sc.Plot();
            return wb;
        }
    }
}
