using System;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>Copies cells, row/column sizes and merged regions, also between XLS and XLSX where NPOI's own CopyTo cannot.</summary>
    internal static class SheetCopier
    {
        public static void Copy(WorkSheet src, WorkBook dst, string name)
        {
            if (src.Book.Format == dst.Format)
            {
                CopyNative(src, dst.Native, name);
                dst.Touch();
                return;
            }

            var from = src.Native;
            var to = dst.Native.CreateSheet(name);
            var lastCol = 0;
            foreach (var row in from.Cast<XL.IRow>().Where(r => r != null))
            {
                var dr = to.CreateRow(row.RowNum);
                if (row.ZeroHeight) dr.ZeroHeight = true;
                if (row.Height >= 0 && row.Height != from.DefaultRowHeight) dr.Height = row.Height; // XSSFRow.HasCustomHeight is not implemented in NPOI 2.7
                foreach (var c in row)
                {
                    lastCol = Math.Max(lastCol, c.ColumnIndex);
                    var d = dr.CreateCell(c.ColumnIndex);
                    CopyValue(c, d);
                    d.CellStyle = dst.ImportStyle(c.CellStyle, src.Book);
                    if (c.Hyperlink != null)
                    {
                        var link = dst.Native.GetCreationHelper().CreateHyperlink(c.Hyperlink.Type);
                        link.Address = c.Hyperlink.Address;
                        d.Hyperlink = link;
                    }
                    if (c.CellComment != null)
                        dst.Wrap(to).GetCellAt(row.RowNum, c.ColumnIndex).AddComment(c.CellComment.String?.String, c.CellComment.Author, c.CellComment.Visible);
                }
            }
            for (var col = 0; col <= lastCol; col++)
            {
                to.SetColumnWidth(col, Math.Min(from.GetColumnWidth(col), 255 * 256));
                if (from.IsColumnHidden(col)) to.SetColumnHidden(col, true);
            }
            foreach (var m in from.MergedRegions) to.AddMergedRegion(new CellRangeAddress(m.FirstRow, m.LastRow, m.FirstColumn, m.LastColumn));
            dst.Touch();
        }

        /// <summary>
        /// NPOI's copy fails (NullReferenceException) on a picture inserted during this session: its package relationship does not exist until
        /// the file is written. The sheet is then copied from a saved-and-reloaded twin of its workbook, whose pictures are complete.
        /// </summary>
        internal static void CopyNative(WorkSheet src, XL.IWorkbook dst, string name)
        {
            var before = dst.NumberOfSheets;
            try { src.Native.CopyTo(dst, name, true, true); CopyXlsPictures(src.Native, dst.GetSheet(name)); }
            catch (NullReferenceException)
            {
                while (dst.NumberOfSheets > before)   // the failed attempt had already added the empty sheet
                {
                    var last = dst.NumberOfSheets - 1;
                    SheetParts.Release(dst.GetSheetAt(last));
                    dst.RemoveSheetAt(last);
                }
                var twin = WorkBook.Load(src.Book.Serialize(src.Book.Format, null));
                twin.Native.GetSheetAt(src.Index).CopyTo(dst, name, true, true);
            }
        }

        /// <summary>NPOI's copy of an .xls sheet leaves its pictures behind (the .xlsx copy keeps them).</summary>
        static void CopyXlsPictures(XL.ISheet from, XL.ISheet to)
        {
            if (!(from.DrawingPatriarch is HSSFPatriarch patriarch) || !(to is HSSFSheet)) return;
            HSSFPatriarch target = null;
            foreach (var pic in patriarch.Children.OfType<HSSFPicture>())
            {
                var a = pic.ClientAnchor;
                var index = to.Workbook.AddPicture(pic.PictureData.Data, pic.PictureData.PictureType);
                target = target ?? (HSSFPatriarch)to.CreateDrawingPatriarch();
                target.CreatePicture(new HSSFClientAnchor(a.Dx1, a.Dy1, a.Dx2, a.Dy2, a.Col1, a.Row1, a.Col2, a.Row2) { AnchorType = a.AnchorType }, index);
            }
        }

        internal static void CopyValue(XL.ICell s, XL.ICell d)
        {
            switch (s.CellType)
            {
                case XL.CellType.Numeric: d.SetCellValue(s.NumericCellValue); break;
                case XL.CellType.String: d.SetCellValue(s.StringCellValue); break;
                case XL.CellType.Boolean: d.SetCellValue(s.BooleanCellValue); break;
                case XL.CellType.Error: d.SetCellErrorValue(s.ErrorCellValue); break;
                case XL.CellType.Formula: d.SetCellFormula(s.CellFormula); break; // same position in another workbook or a moved cell: the text stays
            }
        }

        /// <summary>Re-creates a style in a workbook of the other file format, attribute by attribute (colors become RGB / nearest palette entry).</summary>
        public static XL.ICellStyle CrossStyle(WorkBook srcBook, XL.ICellStyle s, WorkBook dstBook)
        {
            var sw = srcBook.Native; var dw = dstBook.Native;
            var d = dw.CreateCellStyle();
            d.Alignment = s.Alignment; d.VerticalAlignment = s.VerticalAlignment; d.WrapText = s.WrapText;
            d.Indention = s.Indention; d.Rotation = s.Rotation; d.ShrinkToFit = s.ShrinkToFit;
            d.IsLocked = s.IsLocked; d.IsHidden = s.IsHidden;
            d.DataFormat = dw.CreateDataFormat().GetFormat(s.GetDataFormatString());

            d.BorderLeft = s.BorderLeft; d.BorderRight = s.BorderRight; d.BorderTop = s.BorderTop; d.BorderBottom = s.BorderBottom;
            foreach (var side in new[] { Side.Left, Side.Right, Side.Top, Side.Bottom })
            {
                var hex = Colors.GetBorder(sw, s, side);
                if (hex != null) Colors.SetBorder(dw, d, side, Colors.Parse(hex));
            }

            d.FillPattern = s.FillPattern;
            if (s.FillPattern != XL.FillPattern.NoFill)
            {
                var hex = Colors.Hex(s.FillForegroundColorColor);
                if (hex != null) Colors.SetFill(dw, d, Colors.Parse(hex));
            }

            var sf = s.GetFont(sw);
            var df = dw.CreateFont();
            df.FontName = sf.FontName; df.FontHeightInPoints = sf.FontHeightInPoints;
            df.IsBold = sf.IsBold; df.IsItalic = sf.IsItalic; df.IsStrikeout = sf.IsStrikeout;
            df.Underline = sf.Underline; df.TypeOffset = sf.TypeOffset;
            var fc = Colors.GetFont(sw, sf);
            if (fc != null) Colors.SetFont(dw, df, Colors.Parse(fc));
            d.SetFont(df);
            return d;
        }
    }
}
