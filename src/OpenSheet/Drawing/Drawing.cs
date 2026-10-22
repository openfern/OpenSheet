using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Drawing
{
    public enum AnchorType { MoveAndResize = 0, MoveDontResize = 2, DontMoveAndResize = 3 }

    /// <summary>Where a picture or chart sits: a block of cells from its top-left to its bottom-right corner (zero-based).</summary>
    public class Position
    {
        readonly Func<(int col1, int row1, int col2, int row2, AnchorType type)> _get;
        readonly Action<int, int, int, int, AnchorType> _set;

        Position(Func<(int, int, int, int, AnchorType)> get, Action<int, int, int, int, AnchorType> set) { _get = get; _set = set; }

        internal static Position Of(XL.IClientAnchor a) => new Position(
            () => (a.Col1, a.Row1, a.Col2, a.Row2, (AnchorType)(int)a.AnchorType),
            (c1, r1, c2, r2, t) => { a.Col1 = c1; a.Row1 = r1; a.Col2 = c2; a.Row2 = r2; a.AnchorType = (XL.AnchorType)(int)t; });

        // XLSX: bound to the drawing's own XML so edits are what gets saved.
        internal static Position Of(CT_TwoCellAnchor a) => new Position(
            () => (a.from.col, a.from.row, a.to.col, a.to.row, Type(a.editAs)),
            (c1, r1, c2, r2, t) =>
            {
                a.from.col = c1; a.from.row = r1; a.to.col = c2; a.to.row = r2;
                a.editAs = EditAs(t); a.editAsSpecified = true;
            });

        static AnchorType Type(ST_EditAs e) => e == ST_EditAs.oneCell ? AnchorType.MoveDontResize : e == ST_EditAs.absolute ? AnchorType.DontMoveAndResize : AnchorType.MoveAndResize;
        static ST_EditAs EditAs(AnchorType t) => t == AnchorType.MoveDontResize ? ST_EditAs.oneCell : t == AnchorType.DontMoveAndResize ? ST_EditAs.absolute : ST_EditAs.twoCell;

        void Update(int? c1 = null, int? r1 = null, int? c2 = null, int? r2 = null, AnchorType? t = null)
        {
            var g = _get();
            _set(c1 ?? g.col1, r1 ?? g.row1, c2 ?? g.col2, r2 ?? g.row2, t ?? g.type);
        }

        public int LeftColumnIndex { get => _get().col1; set => Update(c1: value); }
        public string LeftColumnLetter { get => A1.Column(_get().col1); set => Update(c1: A1.Column(value)); }
        public int RightColumnIndex { get => _get().col2; set => Update(c2: value); }
        public string RightColumnLetter { get => A1.Column(_get().col2); set => Update(c2: A1.Column(value)); }
        public int TopRowIndex { get => _get().row1; set => Update(r1: value); }
        public int BottomRowIndex { get => _get().row2; set => Update(r2: value); }
        public AnchorType AnchorType { get => _get().type; set => Update(t: value); }
    }
}

namespace OpenSheet.Drawing.Images
{
    /// <summary>The numbers are Excel's picture type codes.</summary>
    public enum ImageFormat { JPEG = 5, PNG = 6, GIF = 8, TIFF = 9, BMP = 11 }

    public interface IImage
    {
        int Id { get; }
        Position Position { get; }
        ImageFormat ImageFormat { get; }
        byte[] Data { get; }
        /// <summary>Back to the picture's natural size.</summary>
        void Resize();
        void Resize(double scale);
        void Resize(double scaleX, double scaleY);
        /// <summary>
        /// The picture as an ImageSharp image (SixLabors.ImageSharp 2.1, Apache-2.0, which NPOI already brings along). Dispose it when done.
        /// The other graphics type the reference offers, <c>IronSoftware.Drawing.AnyBitmap</c>, belongs to a proprietary package and has no counterpart.
        /// </summary>
        SixLabors.ImageSharp.Image ToImage();
    }

    internal static class ImageKit
    {
        /// <summary>Picture format from the file's first bytes; null if it is not one of the supported formats.</summary>
        public static ImageFormat? Sniff(byte[] d)
        {
            if (d == null || d.Length < 4) return null;
            if (d[0] == 0x89 && d[1] == 0x50 && d[2] == 0x4E && d[3] == 0x47) return ImageFormat.PNG;
            if (d[0] == 0xFF && d[1] == 0xD8) return ImageFormat.JPEG;
            if (d[0] == 0x47 && d[1] == 0x49 && d[2] == 0x46) return ImageFormat.GIF;
            if (d[0] == 0x42 && d[1] == 0x4D) return ImageFormat.BMP;
            if ((d[0] == 0x49 && d[1] == 0x49 && d[2] == 0x2A) || (d[0] == 0x4D && d[1] == 0x4D && d[3] == 0x2A)) return ImageFormat.TIFF;
            return null;
        }

        static int BE16(byte[] d, int i) => (d[i] << 8) | d[i + 1];
        static int LE16(byte[] d, int i) => d[i] | (d[i + 1] << 8);
        static int LE32(byte[] d, int i) => d[i] | (d[i + 1] << 8) | (d[i + 2] << 16) | (d[i + 3] << 24);
        static int BE32(byte[] d, int i) => (d[i] << 24) | (d[i + 1] << 16) | (d[i + 2] << 8) | d[i + 3];

        /// <summary>Pixel size read from the file header, or null if it cannot be determined.</summary>
        public static (int w, int h)? Size(byte[] d)
        {
            try
            {
                switch (Sniff(d))
                {
                    case ImageFormat.PNG: return (BE32(d, 16), BE32(d, 20));
                    case ImageFormat.GIF: return (LE16(d, 6), LE16(d, 8));
                    case ImageFormat.BMP: return (Math.Abs(LE32(d, 18)), Math.Abs(LE32(d, 22)));
                    case ImageFormat.JPEG:
                        for (var i = 2; i + 9 < d.Length;)
                        {
                            if (d[i] != 0xFF) { i++; continue; }
                            var m = d[i + 1];
                            if (m >= 0xC0 && m <= 0xCF && m != 0xC4 && m != 0xC8 && m != 0xCC) return (BE16(d, i + 7), BE16(d, i + 5));
                            i += 2 + BE16(d, i + 2);
                        }
                        return null;
                    case ImageFormat.TIFF: return Tiff(d);
                    default: return null;
                }
            }
            catch (IndexOutOfRangeException) { return null; }
        }

        static (int, int)? Tiff(byte[] d)
        {
            var le = d[0] == 0x49;
            int U16(int i) => le ? LE16(d, i) : BE16(d, i);
            int U32(int i) => le ? LE32(d, i) : BE32(d, i);
            int w = 0, h = 0;
            var ifd = U32(4);
            for (var e = 0; e < U16(ifd); e++)
            {
                var p = ifd + 2 + e * 12;
                var tag = U16(p);
                var v = U16(p + 2) == 3 ? U16(p + 8) : U32(p + 8);
                if (tag == 256) w = v; else if (tag == 257) h = v;
            }
            return w > 0 && h > 0 ? ((int, int)?)(w, h) : null;
        }

        const double EmuPerPixel = 9525;

        /// <summary>Bottom-right corner (cell + offset in EMU) of a picture of the given pixel size placed at <paramref name="from"/>.</summary>
        public static void EndFor(XL.ISheet sheet, CT_Marker from, double widthPx, double heightPx, CT_Marker to)
        {
            var remaining = widthPx * EmuPerPixel; var col = from.col; var off = (double)from.colOff;
            for (; col < 16383; col++)
            {
                var avail = sheet.GetColumnWidthInPixels(col) * EmuPerPixel - off;
                if (remaining <= avail) break;
                remaining -= avail; off = 0;
            }
            to.col = col; to.colOff = (long)(off + remaining);

            remaining = heightPx * EmuPerPixel; var row = from.row; off = from.rowOff;
            for (; row < 1048575; row++)
            {
                var pts = sheet.GetRow(row)?.HeightInPoints ?? sheet.DefaultRowHeightInPoints;
                var avail = pts * 96.0 / 72.0 * EmuPerPixel - off;
                if (remaining <= avail) break;
                remaining -= avail; off = 0;
            }
            to.row = row; to.rowOff = (long)(off + remaining);
        }
    }

    /// <summary>A picture of an XLS sheet, through NPOI's picture object.</summary>
    internal sealed class XlsImage : IImage
    {
        readonly HSSFPicture _pic;
        public XlsImage(int id, HSSFPicture pic) { Id = id; _pic = pic; }

        public int Id { get; }
        public Position Position => Position.Of(_pic.ClientAnchor);
        public ImageFormat ImageFormat => (ImageFormat)(int)_pic.PictureData.PictureType;
        public byte[] Data => _pic.PictureData.Data;
        public SixLabors.ImageSharp.Image ToImage() => SixLabors.ImageSharp.Image.Load(Data);
        public void Resize() => _pic.Resize();
        public void Resize(double scale) => _pic.Resize(scale);
        public void Resize(double scaleX, double scaleY) => _pic.Resize(scaleX, scaleY);
    }

    /// <summary>
    /// A picture of an XLSX sheet, read straight from the drawing XML. (NPOI 2.7's shape list fails on drawings that
    /// contain a chart, so it is not used.)
    /// </summary>
    internal sealed class XlsxImage : IImage
    {
        readonly XSSFSheet _sheet;
        readonly XSSFDrawing _drawing;
        readonly CT_TwoCellAnchor _anchor;

        public XlsxImage(int id, XSSFSheet sheet, XSSFDrawing drawing, CT_TwoCellAnchor anchor)
        {
            Id = id; _sheet = sheet; _drawing = drawing; _anchor = anchor;
        }

        public int Id { get; }
        public Position Position => Position.Of(_anchor);

        XSSFPictureData Part => (XSSFPictureData)_drawing.GetRelationById(_anchor.picture.blipFill.blip.embed);
        public byte[] Data => Part.Data;
        public SixLabors.ImageSharp.Image ToImage() => SixLabors.ImageSharp.Image.Load(Data);
        public ImageFormat ImageFormat => ImageKit.Sniff(Data) ?? (ImageFormat)(int)Part.PictureType;

        public void Resize() => Resize(1, 1);
        public void Resize(double scale) => Resize(scale, scale);

        public void Resize(double scaleX, double scaleY)
        {
            if (scaleX <= 0 || scaleY <= 0) throw new ArgumentOutOfRangeException(nameof(scaleX), "Scale must be above zero.");
            var size = ImageKit.Size(Data) ?? throw new NotSupportedException("The size of this picture could not be read.");
            ImageKit.EndFor(_sheet, _anchor.from, size.w * scaleX, size.h * scaleY, _anchor.to);
        }

        public static IEnumerable<CT_TwoCellAnchor> Anchors(XSSFDrawing d) =>
            d.GetCTDrawing().CellAnchors.OfType<CT_TwoCellAnchor>().Where(a => a.picture != null);
    }
}
