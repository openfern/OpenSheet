using System;
using OpenSheet.Drawing;
using OpenSheet.Styles;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>
    /// Live view of the style of a range. Reads come from the top-left cell; writes apply to every cell,
    /// creating cells that do not exist yet (Excel also stores formatting-only cells).
    /// </summary>
    internal sealed class RangeStyle : IStyle
    {
        readonly Range _r;
        public RangeStyle(Range range) { _r = range; }

        WorkBook Book => _r.Sheet.Book;
        internal XL.IWorkbook Wb => Book.Native;
        internal XL.ICellStyle Probe() => _r.PeekFirst()?.CellStyle ?? Wb.GetCellStyleAt(0);

        internal void Edit(string key, Action<XL.ICellStyle> mutate)
        {
            foreach (var c in _r.EnsureCells()) Book.Styles.Apply(c, key, mutate);
            Book.Touch();
        }

        internal void EditFont(string key, Action<XL.IFont> mutate)
        {
            foreach (var c in _r.EnsureCells()) Book.Styles.ApplyFont(c, key, mutate);
            Book.Touch();
        }

        internal XL.IFont ProbeFont() => Probe().GetFont(Wb);

        public IFont Font => new RangeFont(this);
        public IBorder LeftBorder => new RangeBorder(this, Side.Left);
        public IBorder RightBorder => new RangeBorder(this, Side.Right);
        public IBorder TopBorder => new RangeBorder(this, Side.Top);
        public IBorder BottomBorder => new RangeBorder(this, Side.Bottom);
        public IBorder DiagonalBorder => new RangeBorder(this, Side.Diagonal);

        public string BackgroundColor
        {
            get
            {
                var s = Probe();
                return s.FillPattern == XL.FillPattern.NoFill ? "#FFFFFF" : Colors.Hex(s.FillForegroundColorColor) ?? "#FFFFFF";
            }
            set => SetBackgroundColor(value);
        }

        public void SetBackgroundColor(string rgb)
        {
            var bytes = Colors.Parse(rgb);
            var hex = Colors.Hex(new RgbColor(bytes));
            Edit("bg:" + hex, s =>
            {
                Colors.SetFill(Wb, s, bytes);
                if (s.FillPattern == XL.FillPattern.NoFill) s.FillPattern = XL.FillPattern.SolidForeground;
            });
        }

        public void SetBackgroundColor(Color color) => SetBackgroundColor(color.ToHex());

        public DiagonalBorderDirection DiagonalBorderDirection
        {
            get => (DiagonalBorderDirection)(int)Probe().BorderDiagonal;
            set { var v = (XL.BorderDiagonal)(int)value; Edit("diag:" + value, s => Colors.SetDiagonalDirection(Wb, s, v)); }
        }

        public HorizontalAlignment HorizontalAlignment
        {
            get => (HorizontalAlignment)(int)Probe().Alignment;
            set { var v = (XL.HorizontalAlignment)(int)value; Edit("ha:" + value, s => s.Alignment = v); }
        }

        public VerticalAlignment VerticalAlignment
        {
            get => (VerticalAlignment)(int)Probe().VerticalAlignment;
            set { var v = (XL.VerticalAlignment)(int)value; Edit("va:" + value, s => s.VerticalAlignment = v); }
        }

        public FillPattern FillPattern
        {
            get => (FillPattern)(int)Probe().FillPattern;
            set { var v = (XL.FillPattern)(int)value; Edit("fp:" + value, s => s.FillPattern = v); }
        }

        public short Rotation { get => Probe().Rotation; set => Edit("rot:" + value, s => s.Rotation = value); }
        public short Indention { get => Probe().Indention; set => Edit("ind:" + value, s => s.Indention = value); }
        public bool WrapText { get => Probe().WrapText; set => Edit("wrap:" + value, s => s.WrapText = value); }
        public bool ShrinkToFit { get => Probe().ShrinkToFit; set => Edit("shrink:" + value, s => s.ShrinkToFit = value); }
    }

    /// <summary>Minimal IColor so a parsed color can reuse <see cref="Colors.Hex"/> for normalization.</summary>
    internal sealed class RgbColor : XL.IColor
    {
        public RgbColor(byte[] rgb) { RGB = rgb; }
        public short Indexed => 0;
        public byte[] RGB { get; }
    }

    internal sealed class RangeFont : IFont
    {
        readonly RangeStyle _s;
        public RangeFont(RangeStyle s) { _s = s; }

        public bool Bold { get => _s.ProbeFont().IsBold; set => _s.EditFont("b:" + value, f => f.IsBold = value); }
        public bool Italic { get => _s.ProbeFont().IsItalic; set => _s.EditFont("i:" + value, f => f.IsItalic = value); }
        public bool Strikeout { get => _s.ProbeFont().IsStrikeout; set => _s.EditFont("s:" + value, f => f.IsStrikeout = value); }

        public FontUnderlineType Underline
        {
            get => (FontUnderlineType)(int)_s.ProbeFont().Underline;
            set { var v = (XL.FontUnderlineType)(int)value; _s.EditFont("u:" + value, f => f.Underline = v); }
        }

        public FontScript FontScript
        {
            get => (FontScript)(int)_s.ProbeFont().TypeOffset;
            set { var v = (XL.FontSuperScript)(int)value; _s.EditFont("sc:" + value, f => f.TypeOffset = v); }
        }

        public string Name { get => _s.ProbeFont().FontName; set => _s.EditFont("n:" + value, f => f.FontName = value); }

        public short Height
        {
            get => (short)Math.Round(_s.ProbeFont().FontHeightInPoints);
            set => _s.EditFont("h:" + value, f => f.FontHeightInPoints = value);
        }

        public string Color
        {
            get => Colors.GetFont(_s.Wb, _s.ProbeFont()) ?? "#000000";
            set => SetColor(value);
        }

        public void SetColor(string rgb)
        {
            var bytes = Colors.Parse(rgb);
            _s.EditFont("c:" + Colors.Hex(new RgbColor(bytes)), f => Colors.SetFont(_s.Wb, f, bytes));
        }

        public void SetColor(Color color) => SetColor(color.ToHex());
    }

    internal sealed class RangeBorder : IBorder
    {
        readonly RangeStyle _s;
        readonly Side _side;
        public RangeBorder(RangeStyle s, Side side) { _s = s; _side = side; }

        public BorderType Type
        {
            get
            {
                var st = _s.Probe();
                switch (_side)
                {
                    case Side.Left: return (BorderType)(int)st.BorderLeft;
                    case Side.Right: return (BorderType)(int)st.BorderRight;
                    case Side.Top: return (BorderType)(int)st.BorderTop;
                    case Side.Bottom: return (BorderType)(int)st.BorderBottom;
                    default: return (BorderType)(int)st.BorderDiagonalLineStyle;
                }
            }
            set
            {
                var v = (XL.BorderStyle)(int)value;
                _s.Edit($"bt:{_side}:{value}", st =>
                {
                    switch (_side)
                    {
                        case Side.Left: st.BorderLeft = v; break;
                        case Side.Right: st.BorderRight = v; break;
                        case Side.Top: st.BorderTop = v; break;
                        case Side.Bottom: st.BorderBottom = v; break;
                        default: st.BorderDiagonalLineStyle = v; break;
                    }
                });
            }
        }

        public string Color
        {
            get => Colors.GetBorder(_s.Wb, _s.Probe(), _side) ?? "#000000";
            set => SetColor(value);
        }

        public void SetColor(string rgb)
        {
            var bytes = Colors.Parse(rgb);
            _s.Edit($"bc:{_side}:{Colors.Hex(new RgbColor(bytes))}", st => Colors.SetBorder(_s.Wb, st, _side, bytes));
        }

        public void SetColor(Color color) => SetColor(color.ToHex());

        public void SetBorderStyle(string rgb, BorderType type = BorderType.Thin) { Type = type; SetColor(rgb); }
        public void SetBorderStyle(Color color, BorderType type = BorderType.Thin) { Type = type; SetColor(color); }
    }
}
