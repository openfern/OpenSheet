using OpenSheet.Drawing;

namespace OpenSheet.Styles
{
    public enum BorderType
    {
        None = 0, Thin = 1, Medium = 2, Dashed = 3, Dotted = 4, Thick = 5, Double = 6, Hair = 7,
        MediumDashed = 8, DashDot = 9, MediumDashDot = 10, DashDotDot = 11, MediumDashDotDot = 12, SlantedDashDot = 13,
    }

    public enum DiagonalBorderDirection { None = 0, Backward = 1, Forward = 2, Both = 3 }

    public enum FillPattern
    {
        NoFill = 0, SolidForeground = 1, FineDots = 2, AltBars = 3, SparseDots = 4, ThickHorizontalBands = 5,
        ThickVerticalBands = 6, ThickBackwardDiagonals = 7, ThickForwardDiagonals = 8, BigSpots = 9, Bricks = 10,
        ThinHorizontalBands = 11, ThinVerticalBands = 12, ThinBackwardDiagonals = 13, ThinForwardDiagonals = 14,
        Squares = 15, Diamonds = 16, LessDots = 17, LeastDots = 18,
    }

    public enum FontScript { None = 0, Super = 1, Sub = 2 }

    public enum FontUnderlineType { None = 0, Single = 1, Double = 2, SingleAccounting = 33, DoubleAccounting = 34 }

    public enum HorizontalAlignment { General = 0, Left = 1, Center = 2, Right = 3, Fill = 4, Justify = 5, CenterSelection = 6, Distributed = 7 }

    public enum VerticalAlignment { None = -1, Top = 0, Center = 1, Bottom = 2, Justify = 3, Distributed = 4 }

    /// <summary>One side of a cell border. Colors are "#RRGGBB" strings.</summary>
    public interface IBorder
    {
        BorderType Type { get; set; }
        string Color { get; set; }
        void SetColor(string rgb);
        void SetColor(Color color);
        void SetBorderStyle(string rgb, BorderType type = BorderType.Thin);
        void SetBorderStyle(Color color, BorderType type = BorderType.Thin);
    }

    public interface IFont
    {
        bool Bold { get; set; }
        bool Italic { get; set; }
        bool Strikeout { get; set; }
        FontUnderlineType Underline { get; set; }
        FontScript FontScript { get; set; }
        string Name { get; set; }
        /// <summary>Height in points.</summary>
        short Height { get; set; }
        string Color { get; set; }
        void SetColor(string rgb);
        void SetColor(Color color);
    }

    public interface IStyle
    {
        IFont Font { get; }
        string BackgroundColor { get; set; }
        IBorder LeftBorder { get; }
        IBorder RightBorder { get; }
        IBorder TopBorder { get; }
        IBorder BottomBorder { get; }
        IBorder DiagonalBorder { get; }
        DiagonalBorderDirection DiagonalBorderDirection { get; set; }
        HorizontalAlignment HorizontalAlignment { get; set; }
        VerticalAlignment VerticalAlignment { get; set; }
        FillPattern FillPattern { get; set; }
        short Rotation { get; set; }
        short Indention { get; set; }
        bool WrapText { get; set; }
        bool ShrinkToFit { get; set; }
        void SetBackgroundColor(string rgb);
        void SetBackgroundColor(Color color);
    }

    public interface IStylable
    {
        IStyle Style { get; }
    }
}
