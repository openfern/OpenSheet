using System;
using System.IO;
using OpenSheet.Styles;
using Xunit;

namespace OpenSheet.Tests
{
    /// <summary>Runs the README quick-start so the documented example cannot rot.</summary>
    public class ReadmeTests : TestBase
    {
        [Fact]
        public void QuickStart_Works()
        {
            var prices = Tmp("prices.xlsx");
            var book = WorkBook.Create();
            var sheet = book.DefaultWorkSheet;

            sheet["A1"].Value = "Item"; sheet["B1"].Value = "Price";
            sheet["A2"].Value = "Pen"; sheet["B2"].Value = 1.5;
            sheet["A3"].Value = "Notebook"; sheet["B3"].Value = 4.25;
            sheet["B4"].Formula = "=SUM(B2:B3)";

            sheet["A1:B1"].Style.Font.Bold = true;
            sheet["B2:B4"].FormatString = "#,##0.00";
            sheet["A1:B1"].Style.SetBackgroundColor("#FFE699");

            book.SaveAs(prices);
            book.SaveAs(Tmp("prices.csv"));
            book.SaveAs(Tmp("secret.xlsx"), "password");

            var again = WorkBook.Load(prices);
            decimal total = again.DefaultWorkSheet["B2:B3"].Sum();
            Assert.Equal(5.75m, total);
            Assert.Equal("Item,Price\r\nPen,1.50\r\nNotebook,4.25\r\n,5.75\r\n", File.ReadAllText(Tmp("prices.csv")));
            Assert.Equal("Item", WorkBook.Load(Tmp("secret.xlsx"), "password").DefaultWorkSheet["A1"].Value);

            var seen = "";
            foreach (var row in WorkBook.StreamRows(prices, new StreamReadOptions { HasHeaderRow = true }))
                seen += FormattableString.Invariant($"{row["Item"]?.Value}: {row["Price"].Value};");
            Assert.Equal("Pen: 1.5;Notebook: 4.25;: 5.75;", seen);
        }
    }
}
