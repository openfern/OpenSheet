using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using OpenSheet.DataValidations;
using OpenSheet.Drawing;
using OpenSheet.Drawing.Charts;
using OpenSheet.Drawing.Images;
using OpenSheet.Formatting;
using OpenSheet.Internal;
using XL = NPOI.SS.UserModel;

namespace OpenSheet
{
    // Rules, pictures and charts.
    public partial class WorkSheet
    {
        public DataValidationsCollection DataValidations => new DataValidationsCollection(this);

        public ISheetConditionalFormatting ConditionalFormatting => new SheetConditionalFormatting(this);

        // ---- pictures ----

        /// <summary>The pictures on the sheet, in the order they were added. Ids are positions in that order.</summary>
        public List<IImage> Images
        {
            get
            {
                var d = Native.DrawingPatriarch;
                if (d is XSSFDrawing x) return XlsxImage.Anchors(x).Select((a, i) => (IImage)new XlsxImage(i, (XSSFSheet)Native, x, a)).ToList();
                if (d is HSSFPatriarch h) return h.Children.OfType<HSSFPicture>().Select((p, i) => (IImage)new XlsImage(i, p)).ToList();
                return new List<IImage>();
            }
        }

        public IImage InsertImage(string path, string col1, int row1, string col2, int row2) =>
            InsertImage(path, A1.Column(col1), row1, A1.Column(col2), row2);

        public IImage InsertImage(string path, int col1, int row1, int col2, int row2)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("The file name is empty.", nameof(path));
            var bytes = File.ReadAllBytes(path);
            var fmt = ImageKit.Sniff(bytes) ?? throw new NotSupportedException("The file is not a JPEG, PNG, GIF, BMP or TIFF picture.");
            return InsertImage(bytes, fmt, col1, row1, col2, row2);
        }

        public IImage InsertImage(byte[] bytes, ImageFormat pictureType, int col1, int row1, int col2, int row2)
        {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("The picture data is empty.", nameof(bytes));
            if (col2 < col1 || row2 < row1) throw new ArgumentException("The bottom-right corner must not be above or left of the top-left corner.");
            var index = Book.Native.AddPicture(bytes, (XL.PictureType)(int)pictureType);
            var drawing = Native.CreateDrawingPatriarch();
            var anchor = Book.Native.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = col1; anchor.Row1 = row1; anchor.Col2 = col2; anchor.Row2 = row2;
            anchor.AnchorType = XL.AnchorType.MoveAndResize;
            drawing.CreatePicture(anchor, index);
            Book.Touch();
            return Images.Last();
        }

        public void RemoveImage(int id)
        {
            var images = Images;
            if (id < 0 || id >= images.Count) throw new ArgumentOutOfRangeException(nameof(id), "No picture has that id.");
            if (Native.DrawingPatriarch is XSSFDrawing x)
            {
                var anchor = XlsxImage.Anchors(x).ElementAt(id);
                x.GetCTDrawing().CellAnchors.Remove(anchor);
            }
            else ((HSSFPatriarch)Native.DrawingPatriarch).RemoveShape(((HSSFPatriarch)Native.DrawingPatriarch).Children.OfType<HSSFPicture>().ElementAt(id));
            Book.Touch();
        }
        // ---- charts (XLSX only) ----

        readonly List<SheetChart> _charts = new List<SheetChart>();

        XSSFSheet ChartSheet => Native as XSSFSheet ?? throw new NotSupportedException("Charts need the XLSX format.");

        /// <summary>Charts created in this session first, then those read from the file.</summary>
        public List<IChart> Charts
        {
            get
            {
                var list = new List<IChart>(_charts);
                if (Native is XSSFSheet x && x.DrawingPatriarch is XSSFDrawing d)
                    foreach (var c in d.GetCharts()) if (!_charts.Any(m => ReferenceEquals(m.Native, c))) list.Add(new LoadedChart(c, d));
                return list;
            }
        }

        public IChart CreateChart(ChartType type, int row1, string col1, int row2, string col2) => CreateChart(type, row1, A1.Column(col1), row2, A1.Column(col2));

        public IChart CreateChart(ChartType type, int row1, int col1, int row2, int col2)
        {
            if (!Enum.IsDefined(typeof(ChartType), type)) throw new ArgumentOutOfRangeException(nameof(type));
            if (row2 < row1 || col2 < col1) throw new ArgumentException("The bottom-right corner must not be above or left of the top-left corner.");
            var sheet = ChartSheet;
            var drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            var anchor = drawing.CreateAnchor(0, 0, 0, 0, col1, row1, col2, row2);
            var chart = (XSSFChart)drawing.CreateChart(anchor);
            var ct = drawing.GetCTDrawing().CellAnchors.OfType<NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor>().Last(a => a.graphicFrame != null);
            var created = new SheetChart(this, chart, type, Position.Of(ct));
            _charts.Add(created);
            Book.Touch();
            return created;
        }

        public void RemoveChart(IChart chart)
        {
            if (chart == null) throw new ArgumentNullException(nameof(chart));
            if (chart is LoadedChart loaded) { loaded.Remove(); Book.Touch(); return; }
            var native = (chart as SheetChart)?.Native ?? throw new ArgumentException("That chart does not belong to this library.", nameof(chart));
            var drawing = (XSSFDrawing)ChartSheet.DrawingPatriarch;
            drawing.RemoveChart(native);
            _charts.RemoveAll(c => ReferenceEquals(c.Native, native));
            Book.Touch();
        }

        /// <summary>Called when saving: a chart that got series but was never plotted is plotted now.</summary>
        internal void PlotPendingCharts()
        {
            foreach (var c in _charts) if (c.Dirty) try { c.Plot(); } catch (InvalidOperationException) { /* no series: leave the empty chart */ }
        }
    }
}
