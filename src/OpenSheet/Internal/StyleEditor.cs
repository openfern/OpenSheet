using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    internal enum Side { Left, Right, Top, Bottom, Diagonal }

    /// <summary>Color conversion between "#RRGGBB" strings and the two NPOI color models.</summary>
    internal static class Colors
    {
        public static byte[] Parse(string rgb)
        {
            var c = Drawing.Color.FromHex(rgb);
            return new[] { c.R, c.G, c.B };
        }

        /// <summary>XLSX colors are stored as ARGB (4 bytes); a 3-byte value produces a file that fails schema validation.</summary>
        public static XSSFColor Xssf(byte[] rgb) => new XSSFColor(new byte[] { 0xFF, rgb[0], rgb[1], rgb[2] }, null);

        public static string Hex(XL.IColor color)
        {
            var rgb = color?.RGB;
            if (rgb == null || rgb.Length < 3) return null;
            var o = rgb.Length == 4 ? 1 : 0; // ARGB -> skip alpha
            return $"#{rgb[o]:X2}{rgb[o + 1]:X2}{rgb[o + 2]:X2}";
        }

        // XLS has a 56 color palette; shortcut: nearest palette entry, no custom palette editing.
        static short Index(XL.IWorkbook wb, byte[] rgb) =>
            ((HSSFWorkbook)wb).GetCustomPalette().FindSimilarColor(rgb[0], rgb[1], rgb[2]).Indexed;

        public static void SetFill(XL.IWorkbook wb, XL.ICellStyle st, byte[] rgb)
        {
            if (st is XSSFCellStyle x) x.SetFillForegroundColor(Xssf(rgb));
            else st.FillForegroundColor = Index(wb, rgb);
        }

        public static void SetBorder(XL.IWorkbook wb, XL.ICellStyle st, Side side, byte[] rgb)
        {
            if (st is XSSFCellStyle x)
            {
                var c = Xssf(rgb);
                switch (side)
                {
                    case Side.Left: x.SetLeftBorderColor(c); break;
                    case Side.Right: x.SetRightBorderColor(c); break;
                    case Side.Top: x.SetTopBorderColor(c); break;
                    case Side.Bottom: x.SetBottomBorderColor(c); break;
                    default: x.SetDiagonalBorderColor(c); break;
                }
                return;
            }
            var i = Index(wb, rgb);
            switch (side)
            {
                case Side.Left: st.LeftBorderColor = i; break;
                case Side.Right: st.RightBorderColor = i; break;
                case Side.Top: st.TopBorderColor = i; break;
                case Side.Bottom: st.BottomBorderColor = i; break;
                default: st.BorderDiagonalColor = i; break;
            }
        }

        public static void SetFont(XL.IWorkbook wb, XL.IFont f, byte[] rgb)
        {
            if (f is XSSFFont x) x.SetColor(Xssf(rgb));
            else f.Color = Index(wb, rgb);
        }

        public static string GetBorder(XL.IWorkbook wb, XL.ICellStyle st, Side side)
        {
            if (st is XSSFCellStyle x)
            {
                switch (side)
                {
                    case Side.Left: return Hex(x.LeftBorderXSSFColor);
                    case Side.Right: return Hex(x.RightBorderXSSFColor);
                    case Side.Top: return Hex(x.TopBorderXSSFColor);
                    case Side.Bottom: return Hex(x.BottomBorderXSSFColor);
                    default: return Hex(x.DiagonalBorderXSSFColor);
                }
            }
            short i;
            switch (side)
            {
                case Side.Left: i = st.LeftBorderColor; break;
                case Side.Right: i = st.RightBorderColor; break;
                case Side.Top: i = st.TopBorderColor; break;
                case Side.Bottom: i = st.BottomBorderColor; break;
                default: i = st.BorderDiagonalColor; break;
            }
            return Hex(((HSSFWorkbook)wb).GetCustomPalette().GetColor(i));
        }

        /// <summary>
        /// Which diagonals are drawn. NPOI's own setter edits whatever border object it is handed and never stores it: with no border applied
        /// that is a throwaway (the direction is silently lost), with one applied it is the border shared by every style using it (their
        /// direction changes too). So copy the border, set the two flags, store it and point this style at it.
        /// </summary>
        public static void SetDiagonalDirection(XL.IWorkbook wb, XL.ICellStyle st, XL.BorderDiagonal direction)
        {
            if (!(st is XSSFCellStyle x)) { st.BorderDiagonal = direction; return; }
            var table = ((XSSFWorkbook)wb).GetStylesSource();
            var xf = x.GetCoreXf();
            NPOI.OpenXmlFormats.Spreadsheet.CT_Border ct;
            if (xf.applyBorder) ct = table.GetBorderAt((int)xf.borderId).GetCTBorder().Copy();
            else
            {
                ct = new NPOI.OpenXmlFormats.Spreadsheet.CT_Border();
                ct.AddNewLeft(); ct.AddNewRight(); ct.AddNewTop(); ct.AddNewBottom(); ct.AddNewDiagonal();
            }
            ct.diagonalUp = ct.diagonalUpSpecified = direction == XL.BorderDiagonal.Forward || direction == XL.BorderDiagonal.Both;
            ct.diagonalDown = ct.diagonalDownSpecified = direction == XL.BorderDiagonal.Backward || direction == XL.BorderDiagonal.Both;
            xf.borderId = (uint)table.PutBorder(new NPOI.XSSF.UserModel.Extensions.XSSFCellBorder(ct));
            xf.applyBorder = true;
        }

        public static string GetFont(XL.IWorkbook wb, XL.IFont f)
        {
            if (f is XSSFFont x) return Hex(x.GetXSSFColor());
            return Hex(((HSSFFont)f).GetHSSFColor((HSSFWorkbook)wb));
        }
    }

    /// <summary>
    /// Applies one style change to a cell. A cell's style is immutable in the file model (shared by index),
    /// so every change clones it; results are memoized per (source style, change) so a bulk edit creates
    /// one style, not one per cell.
    /// </summary>
    internal sealed class StyleEditor
    {
        readonly XL.IWorkbook _wb;
        readonly Dictionary<(short, string), XL.ICellStyle> _memo = new Dictionary<(short, string), XL.ICellStyle>();

        public StyleEditor(XL.IWorkbook wb) { _wb = wb; }

        public void Apply(XL.ICell cell, string key, Action<XL.ICellStyle> mutate)
        {
            // a bulk edit meets the same source style over and over: when it is known by index, no style object is built for it
            var known = CellValues.StyleIndex(cell);
            if (known >= 0 && _memo.TryGetValue(((short)known, key), out var done)) { cell.CellStyle = done; return; }
            var src = cell.CellStyle;
            var k = (src.Index, key);
            if (!_memo.TryGetValue(k, out var dst))
            {
                dst = _wb.CreateCellStyle();
                dst.CloneStyleFrom(src);
                mutate(dst);
                _memo[k] = dst;
            }
            cell.CellStyle = dst;
        }

        public void ApplyFont(XL.ICell cell, string key, Action<XL.IFont> mutate) =>
            Apply(cell, "font:" + key, s =>
            {
                var f = _wb.CreateFont();
                f.CloneStyleFrom(s.GetFont(_wb));
                mutate(f);
                s.SetFont(f);
            });
    }
}
