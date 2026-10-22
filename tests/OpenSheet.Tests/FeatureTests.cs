using System;
using System.Linq;
using OpenSheet.DataValidations;
using OpenSheet.Drawing.Images;
using OpenSheet.Formatting;
using OpenSheet.Styles;
using Xunit;
using CfOp = OpenSheet.Formatting.Enums.ComparisonOperator;
using ComparisonOperator = OpenSheet.DataValidations.ComparisonOperator;

namespace OpenSheet.Tests
{
    public class ValidationTests : TestBase
    {
        [Theory]
        [InlineData("v.xlsx")]
        [InlineData("v.xls")]
        public void Rules_RoundTrip(string file)
        {
            var wb = WorkBook.Create(file.EndsWith("xls") ? ExcelFileFormat.XLS : ExcelFileFormat.XLSX);
            var ws = wb.DefaultWorkSheet;
            ws.DataValidations.AddStringListRule("A1:A5", new[] { "red", "green", "blue" });
            var ints = ws.DataValidations.AddIntegerRule("B1:B5", ComparisonOperator.Between, 1, 10);
            ints.ErrorAlert = ErrorAlert.Warning;
            ints.ErrorBoxTitle = "Bad"; ints.ErrorBoxText = "1 to 10 only";
            ints.PromptBoxTitle = "Hint"; ints.PromptBoxText = "enter 1-10"; ints.ShowPromptBox = true;
            ws.DataValidations.AddDecimalRule("C1", ComparisonOperator.GreaterThan, 0.5m, null);
            ws.DataValidations.AddDateRule("D1", ComparisonOperator.GreaterThanOrEqual, new DateTime(2024, 1, 1), null);
            ws.DataValidations.AddTextLengthRule("E1", ComparisonOperator.LessThanOrEqual, 8, null);
            ws.DataValidations.AddFormulaListRule("F1", "$A$1:$A$3");
            wb.SaveAs(Tmp(file));

            var v = WorkBook.Load(Tmp(file)).DefaultWorkSheet.DataValidations;
            Assert.Equal(6, v.Count);
            var list = v.GetRuleForCell(WorkBook.Load(Tmp(file)).DefaultWorkSheet.GetCellAt(0, 0));
            Assert.Equal(ConstraintType.List, list.ConstraintType);
            Assert.Equal(ComparisonOperator.None, list.ComparisonOperator);
            Assert.True(list.ShowDropDownList);

            var i = v.Single(x => x.ConstraintType == ConstraintType.WholeNumber);
            Assert.Equal(ComparisonOperator.Between, i.ComparisonOperator);
            Assert.Equal("1", i.Formula1); Assert.Equal("10", i.Formula2);
            Assert.Equal(ErrorAlert.Warning, i.ErrorAlert);
            Assert.Equal("Bad", i.ErrorBoxTitle); Assert.Equal("1 to 10 only", i.ErrorBoxText);
            Assert.Equal("Hint", i.PromptBoxTitle); Assert.Equal("enter 1-10", i.PromptBoxText);
            Assert.Equal("B1:B5", i.ConstrainedRangeAddress.Location);

            Assert.Equal(ComparisonOperator.GreaterThan, v.Single(x => x.ConstraintType == ConstraintType.Decimal).ComparisonOperator);
            Assert.Equal(ComparisonOperator.GreaterThanOrEqual, v.Single(x => x.ConstraintType == ConstraintType.Date).ComparisonOperator);
            Assert.Equal(ComparisonOperator.LessThanOrEqual, v.Single(x => x.ConstraintType == ConstraintType.TextLength).ComparisonOperator);
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void Argument_Checks_And_Removal(ExcelFileFormat fmt)
        {
            var ws = WorkBook.Create(fmt).DefaultWorkSheet;
            Assert.Throws<ArgumentException>(() => ws.DataValidations.AddIntegerRule("A1", ComparisonOperator.Between, 1, null));
            Assert.Throws<ArgumentException>(() => ws.DataValidations.AddIntegerRule("A1", ComparisonOperator.None, 1, null));
            Assert.Throws<ArgumentException>(() => ws.DataValidations.AddStringListRule("A1", new string[0]));
            Assert.Throws<ArgumentException>(() => ws.DataValidations.AddStringListRule("A1", new[] { new string('x', 300) }));

            ws.DataValidations.AddIntegerRule("A1:A3", ComparisonOperator.Equal, 5, null);
            ws.DataValidations.AddIntegerRule("B1:B3", ComparisonOperator.Equal, 6, null);
            Assert.True(ws.DataValidations.RangeAlreadyHasDataValidation(ws["A2"].RangeAddress));
            Assert.False(ws.DataValidations.RangeAlreadyHasDataValidation(ws["Z9"].RangeAddress));
            Assert.Equal(1, ws.DataValidations.GetRulesForRange("B2").Count);

            ws.DataValidations.RemoveAllRulesFromRange("A2");
            Assert.Equal(1, ws.DataValidations.Count);
            ws.DataValidations.RemoveAt(0);
            Assert.Equal(0, ws.DataValidations.Count);
        }
    }

    public class ConditionalFormattingTests : TestBase
    {
        [Theory]
        [InlineData("c.xlsx")]
        [InlineData("c.xls")]
        public void Rule_RoundTrip(string file)
        {
            var wb = WorkBook.Create(file.EndsWith("xls") ? ExcelFileFormat.XLS : ExcelFileFormat.XLSX);
            var ws = wb.DefaultWorkSheet;
            var cf = ws.ConditionalFormatting;
            var rule = cf.CreateConditionalFormattingRule(CfOp.GreaterThan, "5");
            rule.FontFormatting.IsBold = true;
            rule.FontFormatting.FontColor = "#FF0000";
            rule.PatternFormatting.BackgroundColor = "#FFFF00";
            rule.BorderFormatting.BottomBorderType = BorderType.Thin;
            cf.AddConditionalFormatting("A1:A10", rule);
            var byFormula = cf.CreateConditionalFormattingRule("MOD(ROW(),2)=0");
            byFormula.PatternFormatting.BackgroundColor = "#EEEEEE";
            cf.AddConditionalFormatting(new[] { "C1:C5", "E1:E5" }, byFormula);
            wb.SaveAs(Tmp(file));

            var back = WorkBook.Load(Tmp(file)).DefaultWorkSheet.ConditionalFormatting;
            Assert.Equal(2, back.ConditionalFormattingsCount);

            var first = back.GetConditionalFormattingAt(0);
            Assert.Equal("A1:A10", first.GetFormattingRanges().Single().Location);
            var r = first.GetRule(0);
            Assert.Equal(ConditionType.CellValueIs, r.ConditionType);
            Assert.Equal(CfOp.GreaterThan, r.ComparisonOperation);
            Assert.Equal("5", r.Formula1);
            Assert.True(r.FontFormatting.IsBold);
            Assert.Equal("#FF0000", r.FontFormatting.FontColor);
            Assert.Equal("#FFFF00", r.PatternFormatting.BackgroundColor);
            Assert.Equal(BorderType.Thin, r.BorderFormatting.BottomBorderType);

            var second = back.GetConditionalFormattingAt(1);
            Assert.Equal(2, second.GetFormattingRanges().Count);
            Assert.Equal(ConditionType.Formula, second.GetRule(0).ConditionType);

            back.RemoveConditionalFormatting(0);
            Assert.Equal(1, back.ConditionalFormattingsCount);
        }

        [Fact]
        public void ConditionType_Lookup()
        {
            Assert.Same(ConditionType.Formula, ConditionType.ForId(2));
            Assert.Null(ConditionType.ForId(99));
            Assert.Equal(ConditionType.IconSet, ConditionType.ForId((byte)6));
        }
    }

    public class ImageTests : TestBase
    {
        // 1x1 transparent PNG
        static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

        [Theory]
        [InlineData("i.xlsx")]
        [InlineData("i.xls")]
        public void Insert_RoundTrip_Resize_Remove(string file)
        {
            var wb = WorkBook.Create(file.EndsWith("xls") ? ExcelFileFormat.XLS : ExcelFileFormat.XLSX);
            var ws = wb.DefaultWorkSheet;
            var img = ws.InsertImage(Png, ImageFormat.PNG, 1, 1, 4, 5);
            Assert.Equal(0, img.Id);
            Assert.Equal(ImageFormat.PNG, img.ImageFormat);
            ws.InsertImage(Png, ImageFormat.PNG, 6, 6, 8, 8);
            wb.SaveAs(Tmp(file));

            var back = WorkBook.Load(Tmp(file)).DefaultWorkSheet;
            Assert.Equal(2, back.Images.Count);
            var a = back.Images[0];
            Assert.Equal(ImageFormat.PNG, a.ImageFormat);
            Assert.Equal(Png, a.Data);
            Assert.Equal(1, a.Position.LeftColumnIndex);
            Assert.Equal(1, a.Position.TopRowIndex);
            Assert.Equal(4, a.Position.RightColumnIndex);
            Assert.Equal(5, a.Position.BottomRowIndex);
            Assert.Equal("B", a.Position.LeftColumnLetter);

            back.RemoveImage(0);
            Assert.Single(back.Images);
            Assert.Equal(6, back.Images[0].Position.LeftColumnIndex);
        }

        [Fact]
        public void InsertFromPath_SniffsFormat_AndRejectsGarbage()
        {
            System.IO.File.WriteAllBytes(Tmp("p.bin"), Png);   // extension does not matter
            var ws = WorkBook.Create().DefaultWorkSheet;
            Assert.Equal(ImageFormat.PNG, ws.InsertImage(Tmp("p.bin"), 0, 0, 2, 2).ImageFormat);

            System.IO.File.WriteAllText(Tmp("g.png"), "not a picture");
            Assert.Throws<NotSupportedException>(() => ws.InsertImage(Tmp("g.png"), 0, 0, 2, 2));
            Assert.Throws<ArgumentException>(() => ws.InsertImage(Png, ImageFormat.PNG, 5, 5, 2, 2));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.RemoveImage(7));
        }

        [Theory]
        [InlineData(ExcelFileFormat.XLSX)]
        [InlineData(ExcelFileFormat.XLS)]
        public void ToImage_DecodesThePicture(ExcelFileFormat format)
        {
            var ws = WorkBook.Create(format).DefaultWorkSheet;
            var img = ws.InsertImage(Png, ImageFormat.PNG, 1, 1, 3, 3);
            using (var decoded = img.ToImage())
            {
                Assert.Equal(1, decoded.Width);
                Assert.Equal(1, decoded.Height);
            }
        }

        [Fact]
        public void Resize_ChangesAnchor()
        {
            var ws = WorkBook.Create().DefaultWorkSheet;
            var img = ws.InsertImage(Png, ImageFormat.PNG, 1, 1, 3, 3);
            img.Resize(1);   // a 1x1 pixel picture at natural size
            Assert.Equal(1, img.Position.LeftColumnIndex);
            Assert.True(img.Position.RightColumnIndex <= 2);
        }
    }
}
