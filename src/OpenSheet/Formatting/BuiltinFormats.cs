using XL = NPOI.SS.UserModel;

namespace OpenSheet.Formatting
{
    /// <summary>
    /// Excel's built-in number format strings (ECMA-376 §18.8.30). Assign to <c>Range.FormatString</c>.
    /// </summary>
    public static class BuiltinFormats
    {
        static string F(int id) => XL.BuiltinFormats.GetBuiltinFormat(id);

        public static readonly string General = F(0);
        public static readonly string Number0 = F(1);
        public static readonly string Number2 = F(2);
        public static readonly string Thousands0 = F(3);
        public static readonly string Thousands2 = F(4);
        public static readonly string Currency0 = F(5);
        public static readonly string Currency0Red = F(6);
        public static readonly string Currency2 = F(7);
        public static readonly string Currency2Red = F(8);
        public static readonly string Percent = F(9);
        public static readonly string Percent2 = F(10);
        public static readonly string Scientific1 = F(11);
        public static readonly string Scientific2 = F(48);
        public static readonly string Fraction1 = F(12);
        public static readonly string Fraction2 = F(13);
        public static readonly string ShortDate = F(14);
        public static readonly string LongDate1 = F(15);
        public static readonly string LongDate2 = F(16);
        public static readonly string LongDate3 = F(17);
        public static readonly string Time1 = F(18);
        public static readonly string Time2 = F(19);
        public static readonly string Time3 = F(20);
        public static readonly string Time4 = F(21);
        public static readonly string ShortDateAndTime = F(22);
        public static readonly string Accounting0 = F(37);
        public static readonly string Accounting0Red = F(38);
        public static readonly string Accounting2 = F(39);
        public static readonly string Accounting2Red = F(40);
        public static readonly string Duration1 = F(45);
        public static readonly string Duration2 = F(46);
        public static readonly string Duration3 = F(47);
        public static readonly string Text = F(49);
    }
}
