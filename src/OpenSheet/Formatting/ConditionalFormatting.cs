using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal;
using OpenSheet.Styles;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Formatting.Enums
{
    public enum ComparisonOperator
    {
        NoComparison = 0, Between = 1, NotBetween = 2, Equal = 3, NotEqual = 4, GreaterThan = 5, LessThan = 6, GreaterThanOrEqual = 7, LessThanOrEqual = 8,
    }

    /// <summary>Old three-value rule kind, kept so existing code compiles.</summary>
    [Obsolete("Use OpenSheet.Formatting.ConditionType, which covers more kinds of rule.")]
    public enum ConditionType
    {
        None = 0, CellValueIs = 1, Formula = 2,
    }
}

namespace OpenSheet.Formatting
{
    /// <summary>Kind of conditional-formatting rule (cell value, formula, color scale, ...).</summary>
    public class ConditionType
    {
        public static readonly ConditionType CellValueIs = new ConditionType(1, "cellIs");
        public static readonly ConditionType Formula = new ConditionType(2, "expression");
        public static readonly ConditionType ColorScale = new ConditionType(3, "colorScale");
        public static readonly ConditionType DataBar = new ConditionType(4, "dataBar");
        public static readonly ConditionType Filter = new ConditionType(5, "filter");
        public static readonly ConditionType IconSet = new ConditionType(6, "iconSet");

        static readonly ConditionType[] All = { CellValueIs, Formula, ColorScale, DataBar, Filter, IconSet };

        ConditionType(byte id, string type) { Id = id; Type = type; }

        public byte Id { get; set; }
        public string Type { get; set; }

        public static ConditionType ForId(byte id) => All.FirstOrDefault(c => c.Id == id);
        public static ConditionType ForId(int id) => id >= 0 && id <= byte.MaxValue ? ForId((byte)id) : null;

        public override string ToString() => Type;
        public override int GetHashCode() => Id;
        public override bool Equals(object obj) => obj is ConditionType o && o.Id == Id;
    }

    public interface IBorderFormatting
    {
        BorderType BottomBorderType { get; set; }
        BorderType DiagonalBorderType { get; set; }
        BorderType LeftBorderType { get; set; }
        BorderType RightBorderType { get; set; }
        BorderType TopBorderType { get; set; }
        string BottomBorderColor { get; set; }
        string DiagonalBorderColor { get; set; }
        string LeftBorderColor { get; set; }
        string RightBorderColor { get; set; }
        string TopBorderColor { get; set; }
    }

    public interface IFontFormatting
    {
        FontScript EscapementType { get; set; }
        string FontColor { get; set; }
        /// <summary>Height in points (-1 when the rule does not change it).</summary>
        int FontHeight { get; set; }
        FontUnderlineType UnderlineType { get; set; }
        bool IsBold { get; set; }
        bool IsItalic { get; set; }
    }

    public interface IPatternFormatting
    {
        string BackgroundColor { get; set; }
        FillPattern FillPattern { get; set; }
    }

    public interface ISheetConditionalFormatting
    {
        int ConditionalFormattingsCount { get; }
        ConditionalFormatting AddConditionalFormatting(string region, ConditionalFormattingRule rule);
        ConditionalFormatting AddConditionalFormatting(string region, ConditionalFormattingRule rule1, ConditionalFormattingRule rule2);
        ConditionalFormatting AddConditionalFormatting(string region, List<ConditionalFormattingRule> cfRules);
        ConditionalFormatting AddConditionalFormatting(IEnumerable<string> regions, ConditionalFormattingRule rule);
        ConditionalFormatting AddConditionalFormatting(IEnumerable<string> regions, ConditionalFormattingRule rule1, ConditionalFormattingRule rule2);
        ConditionalFormatting AddConditionalFormatting(IEnumerable<string> regions, List<ConditionalFormattingRule> cfRules);
        ConditionalFormatting AddConditionalFormatting(ConditionalFormatting cf);
        ConditionalFormattingRule CreateConditionalFormattingRule(Enums.ComparisonOperator comparisonOperation, string formula1, string formula2);
        ConditionalFormattingRule CreateConditionalFormattingRule(Enums.ComparisonOperator comparisonOperation, string formula);
        ConditionalFormattingRule CreateConditionalFormattingRule(string formula);
        ConditionalFormatting GetConditionalFormattingAt(int index);
        void RemoveConditionalFormatting(int index);
    }

    /// <summary>One rule: when it applies and how matching cells look.</summary>
    public class ConditionalFormattingRule
    {
        readonly XL.IWorkbook _wb;
        internal ConditionalFormattingRule(XL.IConditionalFormattingRule native, XL.IWorkbook wb) { Native = native; _wb = wb; }
        internal XL.IConditionalFormattingRule Native { get; }

        // Reading a format creates an empty one when the rule has none, so there is always something to set.
        public IBorderFormatting BorderFormatting => new CfBorder(Native.BorderFormatting ?? Native.CreateBorderFormatting(), _wb);
        public IFontFormatting FontFormatting => new CfFont(Native.FontFormatting ?? Native.CreateFontFormatting(), _wb);
        public IPatternFormatting PatternFormatting => new CfPattern(Native.PatternFormatting ?? Native.CreatePatternFormatting(), _wb);

        public ConditionType ConditionType => ConditionType.ForId(Native.ConditionType.Id);
        public Enums.ComparisonOperator ComparisonOperation => (Enums.ComparisonOperator)(int)Native.ComparisonOperation;
        public string Formula1 => Native.Formula1;
        public string Formula2 => Native.Formula2;
    }

    /// <summary>A set of rules applied to one or more ranges.</summary>
    public class ConditionalFormatting
    {
        readonly XL.IWorkbook _wb;
        internal ConditionalFormatting(XL.IConditionalFormatting native, XL.IWorkbook wb) { Native = native; _wb = wb; }
        internal XL.IConditionalFormatting Native { get; }

        public int NumberOfRules => Native.NumberOfRules;
        public void AddRule(ConditionalFormattingRule cfRule) => Native.AddRule(cfRule.Native);
        public List<RangeAddress> GetFormattingRanges() => Native.GetFormattingRanges().Select(RangeAddress.From).ToList();
        public ConditionalFormattingRule GetRule(int idx) => new ConditionalFormattingRule(Native.GetRule(idx), _wb);
        public void SetRule(int idx, ConditionalFormattingRule cfRule) => Native.SetRule(idx, cfRule.Native);
    }

    internal sealed class SheetConditionalFormatting : ISheetConditionalFormatting
    {
        readonly WorkSheet _s;
        public SheetConditionalFormatting(WorkSheet s) { _s = s; }

        XL.ISheetConditionalFormatting Cf => _s.Native.SheetConditionalFormatting;
        ConditionalFormatting Wrap(int index) => new ConditionalFormatting(Cf.GetConditionalFormattingAt(index), _s.Book.Native);
        static NPOI.SS.Util.CellRangeAddress[] Regions(IEnumerable<string> r) => r.Select(x => A1.Parse(x).ToNpoi()).ToArray();

        public int ConditionalFormattingsCount => Cf.NumConditionalFormattings;

        public ConditionalFormatting AddConditionalFormatting(string region, ConditionalFormattingRule rule) => AddConditionalFormatting(new[] { region }, rule);
        public ConditionalFormatting AddConditionalFormatting(string region, ConditionalFormattingRule r1, ConditionalFormattingRule r2) => AddConditionalFormatting(new[] { region }, r1, r2);
        public ConditionalFormatting AddConditionalFormatting(string region, List<ConditionalFormattingRule> rules) => AddConditionalFormatting(new[] { region }, rules);

        public ConditionalFormatting AddConditionalFormatting(IEnumerable<string> regions, ConditionalFormattingRule rule) =>
            Wrap(Cf.AddConditionalFormatting(Regions(regions), rule.Native));

        public ConditionalFormatting AddConditionalFormatting(IEnumerable<string> regions, ConditionalFormattingRule r1, ConditionalFormattingRule r2) =>
            Wrap(Cf.AddConditionalFormatting(Regions(regions), r1.Native, r2.Native));

        public ConditionalFormatting AddConditionalFormatting(IEnumerable<string> regions, List<ConditionalFormattingRule> rules)
        {
            if (rules == null || rules.Count == 0) throw new ArgumentException("Give at least one rule.", nameof(rules));
            return Wrap(Cf.AddConditionalFormatting(Regions(regions), rules.Select(r => r.Native).ToArray()));
        }

        public ConditionalFormatting AddConditionalFormatting(ConditionalFormatting cf) => Wrap(Cf.AddConditionalFormatting(cf.Native));

        public ConditionalFormattingRule CreateConditionalFormattingRule(Enums.ComparisonOperator op, string formula1, string formula2) =>
            new ConditionalFormattingRule(Cf.CreateConditionalFormattingRule((XL.ComparisonOperator)(int)op, formula1, formula2), _s.Book.Native);

        public ConditionalFormattingRule CreateConditionalFormattingRule(Enums.ComparisonOperator op, string formula) =>
            new ConditionalFormattingRule(Cf.CreateConditionalFormattingRule((XL.ComparisonOperator)(int)op, formula), _s.Book.Native);

        public ConditionalFormattingRule CreateConditionalFormattingRule(string formula) =>
            new ConditionalFormattingRule(Cf.CreateConditionalFormattingRule(formula), _s.Book.Native);

        public ConditionalFormatting GetConditionalFormattingAt(int index) => Wrap(index);
        public void RemoveConditionalFormatting(int index) => Cf.RemoveConditionalFormatting(index);
    }

    // ---- formatting parts: NPOI colors are IColor objects, so one conversion serves XLS and XLSX ----

    static class CfColor
    {
        public static XL.IColor Make(XL.IWorkbook wb, string rgb)
        {
            var b = Colors.Parse(rgb);
            if (wb is XSSFWorkbook) return Colors.Xssf(b);
            return ((HSSFWorkbook)wb).GetCustomPalette().FindSimilarColor(b[0], b[1], b[2]);
        }
    }

    sealed class CfBorder : IBorderFormatting
    {
        readonly XL.IBorderFormatting _b; readonly XL.IWorkbook _wb;
        public CfBorder(XL.IBorderFormatting b, XL.IWorkbook wb) { _b = b; _wb = wb; }

        public BorderType BottomBorderType { get => (BorderType)(int)_b.BorderBottom; set => _b.BorderBottom = (XL.BorderStyle)(int)value; }
        public BorderType DiagonalBorderType { get => (BorderType)(int)_b.BorderDiagonal; set => _b.BorderDiagonal = (XL.BorderStyle)(int)value; }
        public BorderType LeftBorderType { get => (BorderType)(int)_b.BorderLeft; set => _b.BorderLeft = (XL.BorderStyle)(int)value; }
        public BorderType RightBorderType { get => (BorderType)(int)_b.BorderRight; set => _b.BorderRight = (XL.BorderStyle)(int)value; }
        public BorderType TopBorderType { get => (BorderType)(int)_b.BorderTop; set => _b.BorderTop = (XL.BorderStyle)(int)value; }
        public string BottomBorderColor { get => Colors.Hex(_b.BottomBorderColorColor); set => _b.BottomBorderColorColor = CfColor.Make(_wb, value); }
        public string DiagonalBorderColor { get => Colors.Hex(_b.DiagonalBorderColorColor); set => _b.DiagonalBorderColorColor = CfColor.Make(_wb, value); }
        public string LeftBorderColor { get => Colors.Hex(_b.LeftBorderColorColor); set => _b.LeftBorderColorColor = CfColor.Make(_wb, value); }
        public string RightBorderColor { get => Colors.Hex(_b.RightBorderColorColor); set => _b.RightBorderColorColor = CfColor.Make(_wb, value); }
        public string TopBorderColor { get => Colors.Hex(_b.TopBorderColorColor); set => _b.TopBorderColorColor = CfColor.Make(_wb, value); }
    }

    sealed class CfFont : IFontFormatting
    {
        readonly XL.IFontFormatting _f; readonly XL.IWorkbook _wb;
        public CfFont(XL.IFontFormatting f, XL.IWorkbook wb) { _f = f; _wb = wb; }

        public FontScript EscapementType { get => (FontScript)(int)_f.EscapementType; set => _f.EscapementType = (XL.FontSuperScript)(int)value; }
        public string FontColor { get => Colors.Hex(_f.FontColor); set => _f.FontColor = CfColor.Make(_wb, value); }
        public int FontHeight { get => _f.FontHeight < 0 ? -1 : _f.FontHeight / 20; set => _f.FontHeight = value < 0 ? -1 : value * 20; }
        public FontUnderlineType UnderlineType { get => (FontUnderlineType)(int)_f.UnderlineType; set => _f.UnderlineType = (XL.FontUnderlineType)(int)value; }
        public bool IsBold { get => _f.IsBold; set => _f.SetFontStyle(_f.IsItalic, value); }
        public bool IsItalic { get => _f.IsItalic; set => _f.SetFontStyle(value, _f.IsBold); }
    }

    sealed class CfPattern : IPatternFormatting
    {
        readonly XL.IPatternFormatting _p; readonly XL.IWorkbook _wb;
        public CfPattern(XL.IPatternFormatting p, XL.IWorkbook wb) { _p = p; _wb = wb; }

        public string BackgroundColor
        {
            get => Colors.Hex(_p.FillBackgroundColorColor);
            set
            {
                // In a conditional format a solid fill is drawn with the *background* color slot; set both so any reader agrees.
                var c = CfColor.Make(_wb, value);
                _p.FillBackgroundColorColor = c;
                _p.FillForegroundColorColor = c;
                if (_p.FillPattern == XL.FillPattern.NoFill) _p.FillPattern = XL.FillPattern.SolidForeground;
            }
        }

        public FillPattern FillPattern { get => (FillPattern)(int)_p.FillPattern; set => _p.FillPattern = (XL.FillPattern)(int)value; }
    }
}
