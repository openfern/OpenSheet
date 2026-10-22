namespace OpenSheet.Styles
{
    /// <summary>Built-in Excel table style, identified by its name (e.g. "TableStyleMedium2").</summary>
    public class TableStyle
    {
        public TableStyle() { }
        TableStyle(string name) { Name = name; }

        public static readonly TableStyle None = new TableStyle(null);
        public static readonly TableStyle TableStyleLight1 = new TableStyle("TableStyleLight1");
        public static readonly TableStyle TableStyleLight2 = new TableStyle("TableStyleLight2");
        public static readonly TableStyle TableStyleLight3 = new TableStyle("TableStyleLight3");
        public static readonly TableStyle TableStyleLight4 = new TableStyle("TableStyleLight4");
        public static readonly TableStyle TableStyleLight5 = new TableStyle("TableStyleLight5");
        public static readonly TableStyle TableStyleLight6 = new TableStyle("TableStyleLight6");
        public static readonly TableStyle TableStyleLight7 = new TableStyle("TableStyleLight7");
        public static readonly TableStyle TableStyleLight8 = new TableStyle("TableStyleLight8");
        public static readonly TableStyle TableStyleLight9 = new TableStyle("TableStyleLight9");
        public static readonly TableStyle TableStyleLight10 = new TableStyle("TableStyleLight10");
        public static readonly TableStyle TableStyleLight11 = new TableStyle("TableStyleLight11");
        public static readonly TableStyle TableStyleLight12 = new TableStyle("TableStyleLight12");
        public static readonly TableStyle TableStyleLight13 = new TableStyle("TableStyleLight13");
        public static readonly TableStyle TableStyleLight14 = new TableStyle("TableStyleLight14");
        public static readonly TableStyle TableStyleLight15 = new TableStyle("TableStyleLight15");
        public static readonly TableStyle TableStyleLight16 = new TableStyle("TableStyleLight16");
        public static readonly TableStyle TableStyleLight17 = new TableStyle("TableStyleLight17");
        public static readonly TableStyle TableStyleLight18 = new TableStyle("TableStyleLight18");
        public static readonly TableStyle TableStyleLight19 = new TableStyle("TableStyleLight19");
        public static readonly TableStyle TableStyleLight20 = new TableStyle("TableStyleLight20");
        public static readonly TableStyle TableStyleLight21 = new TableStyle("TableStyleLight21");
        public static readonly TableStyle TableStyleMedium1 = new TableStyle("TableStyleMedium1");
        public static readonly TableStyle TableStyleMedium2 = new TableStyle("TableStyleMedium2");
        public static readonly TableStyle TableStyleMedium3 = new TableStyle("TableStyleMedium3");
        public static readonly TableStyle TableStyleMedium4 = new TableStyle("TableStyleMedium4");
        public static readonly TableStyle TableStyleMedium5 = new TableStyle("TableStyleMedium5");
        public static readonly TableStyle TableStyleMedium6 = new TableStyle("TableStyleMedium6");
        public static readonly TableStyle TableStyleMedium7 = new TableStyle("TableStyleMedium7");
        public static readonly TableStyle TableStyleMedium8 = new TableStyle("TableStyleMedium8");
        public static readonly TableStyle TableStyleMedium9 = new TableStyle("TableStyleMedium9");
        public static readonly TableStyle TableStyleMedium10 = new TableStyle("TableStyleMedium10");
        public static readonly TableStyle TableStyleMedium11 = new TableStyle("TableStyleMedium11");
        public static readonly TableStyle TableStyleMedium12 = new TableStyle("TableStyleMedium12");
        public static readonly TableStyle TableStyleMedium13 = new TableStyle("TableStyleMedium13");
        public static readonly TableStyle TableStyleMedium14 = new TableStyle("TableStyleMedium14");
        public static readonly TableStyle TableStyleMedium15 = new TableStyle("TableStyleMedium15");
        public static readonly TableStyle TableStyleMedium16 = new TableStyle("TableStyleMedium16");
        public static readonly TableStyle TableStyleMedium17 = new TableStyle("TableStyleMedium17");
        public static readonly TableStyle TableStyleMedium18 = new TableStyle("TableStyleMedium18");
        public static readonly TableStyle TableStyleMedium19 = new TableStyle("TableStyleMedium19");
        public static readonly TableStyle TableStyleMedium20 = new TableStyle("TableStyleMedium20");
        public static readonly TableStyle TableStyleMedium21 = new TableStyle("TableStyleMedium21");
        public static readonly TableStyle TableStyleMedium22 = new TableStyle("TableStyleMedium22");
        public static readonly TableStyle TableStyleMedium23 = new TableStyle("TableStyleMedium23");
        public static readonly TableStyle TableStyleMedium24 = new TableStyle("TableStyleMedium24");
        public static readonly TableStyle TableStyleMedium25 = new TableStyle("TableStyleMedium25");
        public static readonly TableStyle TableStyleMedium26 = new TableStyle("TableStyleMedium26");
        public static readonly TableStyle TableStyleMedium27 = new TableStyle("TableStyleMedium27");
        public static readonly TableStyle TableStyleMedium28 = new TableStyle("TableStyleMedium28");
        public static readonly TableStyle TableStyleDark1 = new TableStyle("TableStyleDark1");
        public static readonly TableStyle TableStyleDark2 = new TableStyle("TableStyleDark2");
        public static readonly TableStyle TableStyleDark3 = new TableStyle("TableStyleDark3");
        public static readonly TableStyle TableStyleDark4 = new TableStyle("TableStyleDark4");
        public static readonly TableStyle TableStyleDark5 = new TableStyle("TableStyleDark5");
        public static readonly TableStyle TableStyleDark6 = new TableStyle("TableStyleDark6");
        public static readonly TableStyle TableStyleDark7 = new TableStyle("TableStyleDark7");
        public static readonly TableStyle TableStyleDark8 = new TableStyle("TableStyleDark8");
        public static readonly TableStyle TableStyleDark9 = new TableStyle("TableStyleDark9");
        public static readonly TableStyle TableStyleDark10 = new TableStyle("TableStyleDark10");
        public static readonly TableStyle TableStyleDark11 = new TableStyle("TableStyleDark11");

        public virtual string Name { get; set; }
        public virtual bool? ShowLastColumn { get; set; }
        public virtual bool? ShowFirstColumn { get; set; }
        public virtual bool? ShowRowStripes { get; set; }
        public virtual bool? ShowColumnStripes { get; set; }
    }
}
