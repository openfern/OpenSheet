using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Drawing.Charts
{
    public enum ChartType { Column = 1, Scatter = 2, Line = 3, Pie = 4, Bar = 5, Area = 6 }

    public enum LegendPosition { None = 0, Bottom = 1, Left = 2, Right = 3, Top = 4, TopRight = 5 }

    public interface IChartSeries
    {
        /// <summary>Name shown in the legend; null lets Excel name it "Series1", ...</summary>
        string Title { get; set; }
    }

    /// <summary>A chart on a worksheet. Add series, then call <see cref="Plot"/>.</summary>
    public interface IChart
    {
        Position Position { get; }
        /// <summary>Series for a scatter/line/column... chart: <paramref name="xRange"/> are the categories (or X values), <paramref name="yRange"/> the values.</summary>
        IChartSeries AddSeries(string xRange, string yRange);
        /// <summary>A series with the values only; categories are numbered 1, 2, 3...</summary>
        IChartSeries AddSeries(string values);
        void SetLegendPosition(LegendPosition position);
        /// <summary>Writes the series into the chart. Safe to call again after changes.</summary>
        void Plot();
        void SetTitle(string title);
    }

    sealed class ChartSeries : IChartSeries
    {
        public string Title { get; set; }
        public string XRef, YRef;
        public List<Cell> X, Y;
    }

    /// <summary>A chart created by this library. The chart XML is built from the series on <see cref="Plot"/>.</summary>
    sealed class SheetChart : IChart
    {
        const uint CatId = 507160001, ValId = 507160002;
        readonly WorkSheet _sheet;
        readonly XSSFChart _chart;
        internal XSSFChart Native => _chart;
        readonly ChartType _type;
        readonly List<ChartSeries> _series = new List<ChartSeries>();
        readonly Position _position;
        LegendPosition _legend = LegendPosition.Right;
        bool _dirty = true;

        public SheetChart(WorkSheet sheet, XSSFChart chart, ChartType type, Position position)
        {
            _sheet = sheet; _chart = chart; _type = type; _position = position;
        }

        public Position Position => _position;
        internal bool Dirty => _dirty;

        public void SetTitle(string title) { _chart.SetTitle(title); }

        public void SetLegendPosition(LegendPosition position) { _legend = position; _dirty = true; }

        public IChartSeries AddSeries(string values)
        {
            var (f, cells) = Resolve(values);
            var s = new ChartSeries { YRef = f, Y = cells };
            _series.Add(s); _dirty = true;
            return s;
        }

        public IChartSeries AddSeries(string xRange, string yRange)
        {
            var (xf, xc) = Resolve(xRange);
            var (yf, yc) = Resolve(yRange);
            if (xc.Count != yc.Count) throw new ArgumentException($"The ranges differ in size ({xc.Count} and {yc.Count} cells).");
            var s = new ChartSeries { XRef = xf, X = xc, YRef = yf, Y = yc };
            _series.Add(s); _dirty = true;
            return s;
        }

        // A series range is one row or one column; "Sheet2!A1:A9" points at another sheet.
        (string formula, List<Cell> cells) Resolve(string range)
        {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentException("The range is empty.", nameof(range));
            var sheet = _sheet;
            var bang = range.LastIndexOf('!');
            if (bang > 0)
            {
                var name = range.Substring(0, bang).Trim('\'').Replace("''", "'");
                sheet = _sheet.Book.GetWorkSheet(name) ?? throw new ArgumentException($"No worksheet named '{name}'.", nameof(range));
            }
            var a = A1.Parse(range);
            if (a.RowsCount > 1 && a.ColumnsCount > 1) throw new ArgumentException("A series must be a single row or a single column.", nameof(range));
            var cells = new Range(sheet, a).ToArray().ToList();
            return (WorkSheet.Quote(sheet.Name) + "!" + WorkSheet.Absolute(a), cells);
        }

        // ---- building the chart XML (ECMA-376 §21.2) ----

        public void Plot()
        {
            if (_series.Count == 0) throw new InvalidOperationException("Add at least one series before plotting.");
            var ch = _chart.GetCTChart();
            var pa = ch.plotArea ?? (ch.plotArea = new CT_PlotArea());
            // NPOI leaves these lists null until something is added
            pa.barChart = new List<CT_BarChart>(); pa.lineChart = new List<CT_LineChart>(); pa.areaChart = new List<CT_AreaChart>();
            pa.pieChart = new List<CT_PieChart>(); pa.scatterChart = new List<CT_ScatterChart>();
            pa.catAx = new List<CT_CatAx>(); pa.valAx = new List<CT_ValAx>();

            switch (_type)
            {
                case ChartType.Column: case ChartType.Bar: PlotBar(pa, _type == ChartType.Bar); break;
                case ChartType.Line: PlotLine(pa); break;
                case ChartType.Area: PlotArea(pa); break;
                case ChartType.Pie: PlotPie(pa); break;
                case ChartType.Scatter: PlotScatter(pa); break;
                default: throw new NotSupportedException();
            }

            ch.legend = _legend == LegendPosition.None ? null : new CT_Legend
            {
                legendPos = new CT_LegendPos { val = LegendPos(_legend) },
                legendPosSpecified = true,
                overlay = new CT_Boolean { val = 0 },
            };
            ch.plotVisOnly = new CT_Boolean { val = 1 };
            _dirty = false;
        }

        static ST_LegendPos LegendPos(LegendPosition p)
        {
            switch (p)
            {
                case LegendPosition.Bottom: return ST_LegendPos.b;
                case LegendPosition.Left: return ST_LegendPos.l;
                case LegendPosition.Top: return ST_LegendPos.t;
                case LegendPosition.TopRight: return ST_LegendPos.tr;
                default: return ST_LegendPos.r;
            }
        }

        static CT_UnsignedInt U(uint v) => new CT_UnsignedInt { val = v };

        CT_SerTx Tx(ChartSeries s) => s.Title == null ? null : new CT_SerTx { v = s.Title };

        // Categories are text when every cell is text, otherwise numbers.
        static CT_AxDataSource Cat(string f, List<Cell> cells)
        {
            if (cells.All(c => c.IsText || c.IsEmpty) && cells.Any(c => c.IsText))
            {
                var cache = new CT_StrData { ptCount = U((uint)cells.Count), pt = new List<CT_StrVal>() };
                for (var i = 0; i < cells.Count; i++) if (!cells[i].IsEmpty) cache.pt.Add(new CT_StrVal { idx = (uint)i, v = cells[i].StringValue });
                return new CT_AxDataSource { strRef = new CT_StrRef { f = f, strCache = cache } };
            }
            return new CT_AxDataSource { numRef = new CT_NumRef { f = f, numCache = Numbers(cells) } };
        }

        // Without a category range the points are numbered 1..n, written as literals.
        static CT_AxDataSource Numbered(int n)
        {
            var lit = new CT_NumData { formatCode = "General", ptCount = U((uint)n), pt = new List<CT_NumVal>() };
            for (var i = 0; i < n; i++) lit.pt.Add(new CT_NumVal { idx = (uint)i, v = (i + 1).ToString(CultureInfo.InvariantCulture) });
            return new CT_AxDataSource { numLit = lit };
        }

        static CT_NumData Numbers(List<Cell> cells)
        {
            var d = new CT_NumData { formatCode = "General", ptCount = U((uint)cells.Count), pt = new List<CT_NumVal>() };
            for (var i = 0; i < cells.Count; i++)
            {
                var v = cells[i].Read(out var k);
                if (k == Kind.Number || k == Kind.Date) d.pt.Add(new CT_NumVal { idx = (uint)i, v = CellValues.ToDouble(v).ToString("R", CultureInfo.InvariantCulture) });
            }
            return d;
        }

        static CT_NumDataSource Val(string f, List<Cell> cells) => new CT_NumDataSource { numRef = new CT_NumRef { f = f, numCache = Numbers(cells) } };

        CT_AxDataSource CatOf(ChartSeries s) => s.XRef != null ? Cat(s.XRef, s.X) : Numbered(s.Y.Count);

        void Axes(CT_PlotArea pa, ST_AxPos catPos, ST_AxPos valPos)
        {
            pa.catAx.Add(new CT_CatAx
            {
                axId = U(CatId), scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax } },
                delete = new CT_Boolean { val = 0 }, axPos = new CT_AxPos { val = catPos },
                majorTickMark = new CT_TickMark { val = ST_TickMark.@out }, minorTickMark = new CT_TickMark { val = ST_TickMark.none },
                tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo },
                crossAx = U(ValId), crosses = new CT_Crosses { val = ST_Crosses.autoZero },
            });
            pa.valAx.Add(new CT_ValAx
            {
                axId = U(ValId), scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax } },
                delete = new CT_Boolean { val = 0 }, axPos = new CT_AxPos { val = valPos }, majorGridlines = new CT_ChartLines(),
                majorTickMark = new CT_TickMark { val = ST_TickMark.@out }, minorTickMark = new CT_TickMark { val = ST_TickMark.none },
                tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo },
                crossAx = U(CatId), crosses = new CT_Crosses { val = ST_Crosses.autoZero },
                crossBetween = new CT_CrossBetween { val = ST_CrossBetween.between },
            });
        }

        void PlotBar(CT_PlotArea pa, bool horizontal)
        {
            var c = new CT_BarChart
            {
                barDir = new CT_BarDir { val = horizontal ? ST_BarDir.bar : ST_BarDir.col },
                grouping = new CT_BarGrouping { val = ST_BarGrouping.clustered },
                varyColors = new CT_Boolean { val = 0 },
                ser = new List<CT_BarSer>(), axId = new List<CT_UnsignedInt>(),
            };
            for (var i = 0; i < _series.Count; i++)
            {
                var s = _series[i];
                c.ser.Add(new CT_BarSer { idx = U((uint)i), order = U((uint)i), tx = Tx(s), cat = CatOf(s), val = Val(s.YRef, s.Y) });
            }
            c.axId.Add(U(CatId)); c.axId.Add(U(ValId));
            pa.barChart.Add(c);
            Axes(pa, horizontal ? ST_AxPos.l : ST_AxPos.b, horizontal ? ST_AxPos.b : ST_AxPos.l);
        }

        void PlotLine(CT_PlotArea pa)
        {
            var c = new CT_LineChart { grouping = new CT_Grouping { val = ST_Grouping.standard }, varyColors = new CT_Boolean { val = 0 }, marker = new CT_Boolean { val = 1 }, ser = new List<CT_LineSer>(), axId = new List<CT_UnsignedInt>() };
            for (var i = 0; i < _series.Count; i++)
            {
                var s = _series[i];
                // Excel's plain "Line" chart has no point markers; without this per-series switch it draws "Line with markers".
                c.ser.Add(new CT_LineSer { idx = U((uint)i), order = U((uint)i), tx = Tx(s), marker = new NPOI.OpenXmlFormats.Dml.Chart.CT_Marker { symbol = new CT_MarkerStyle { val = ST_MarkerStyle.none } }, cat = CatOf(s), val = Val(s.YRef, s.Y), smooth = new CT_Boolean { val = 0 } });
            }
            c.axId.Add(U(CatId)); c.axId.Add(U(ValId));
            pa.lineChart.Add(c);
            Axes(pa, ST_AxPos.b, ST_AxPos.l);
        }

        void PlotArea(CT_PlotArea pa)
        {
            var c = new CT_AreaChart { grouping = new CT_Grouping { val = ST_Grouping.standard }, varyColors = new CT_Boolean { val = 0 }, ser = new List<CT_AreaSer>(), axId = new List<CT_UnsignedInt>() };
            for (var i = 0; i < _series.Count; i++)
            {
                var s = _series[i];
                c.ser.Add(new CT_AreaSer { idx = U((uint)i), order = U((uint)i), tx = Tx(s), cat = CatOf(s), val = Val(s.YRef, s.Y) });
            }
            c.axId.Add(U(CatId)); c.axId.Add(U(ValId));
            pa.areaChart.Add(c);
            Axes(pa, ST_AxPos.b, ST_AxPos.l);
        }

        void PlotPie(CT_PlotArea pa)
        {
            // firstSliceAng is left out on purpose: its default is 0 and NPOI 2.7 writes it before <ser>, which breaks the schema order
            var c = new CT_PieChart { varyColors = new CT_Boolean { val = 1 }, ser = new List<CT_PieSer>() };
            // a pie shows one series; extra series are ignored by Excel, so only the first is written
            var s = _series[0];
            c.ser.Add(new CT_PieSer { idx = U(0), order = U(0), tx = Tx(s), cat = CatOf(s), val = Val(s.YRef, s.Y) });
            pa.pieChart.Add(c);
        }

        void PlotScatter(CT_PlotArea pa)
        {
            var c = new CT_ScatterChart { scatterStyle = new CT_ScatterStyle { val = ST_ScatterStyle.lineMarker }, varyColors = new CT_Boolean { val = 0 }, ser = new List<CT_ScatterSer>(), axId = new List<CT_UnsignedInt>() };
            for (var i = 0; i < _series.Count; i++)
            {
                var s = _series[i];
                var x = s.XRef != null
                    ? new CT_AxDataSource { numRef = new CT_NumRef { f = s.XRef, numCache = Numbers(s.X) } }
                    : Numbered(s.Y.Count);
                // Excel's "Scatter" chart is markers only: same scatterStyle as the line variants, but the connecting line is switched off.
                var noLine = new NPOI.OpenXmlFormats.Dml.Chart.CT_ShapeProperties { ln = new NPOI.OpenXmlFormats.Dml.CT_LineProperties { noFill = new NPOI.OpenXmlFormats.Dml.CT_NoFillProperties() } };
                c.ser.Add(new CT_ScatterSer { idx = U((uint)i), order = U((uint)i), tx = Tx(s), spPr = noLine, xVal = x, yVal = Val(s.YRef, s.Y), smooth = new CT_Boolean { val = 0 } });
            }
            c.axId.Add(U(CatId)); c.axId.Add(U(ValId));
            pa.scatterChart.Add(c);
            foreach (var (id, cross, pos) in new[] { (CatId, ValId, ST_AxPos.b), (ValId, CatId, ST_AxPos.l) })
                pa.valAx.Add(new CT_ValAx
                {
                    axId = U(id), scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax } },
                    delete = new CT_Boolean { val = 0 }, axPos = new CT_AxPos { val = pos }, majorGridlines = pos == ST_AxPos.l ? new CT_ChartLines() : null,
                    majorTickMark = new CT_TickMark { val = ST_TickMark.@out }, minorTickMark = new CT_TickMark { val = ST_TickMark.none },
                    tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo },
                    crossAx = U(cross), crosses = new CT_Crosses { val = ST_Crosses.autoZero },
                    crossBetween = new CT_CrossBetween { val = ST_CrossBetween.midCat },
                });
        }
    }

    /// <summary>A chart read from a file. Titles and legend can change; its data cannot be edited.</summary>
    sealed class LoadedChart : IChart
    {
        readonly XSSFChart _chart;
        readonly XSSFDrawing _drawing;
        public LoadedChart(XSSFChart chart, XSSFDrawing drawing) { _chart = chart; _drawing = drawing; }

        /// <summary>Where the chart sits. The drawing holds an anchor whose graphic frame points at the chart part by relationship id.</summary>
        public Position Position
        {
            get
            {
                var anchor = Anchor();
                return anchor != null ? Position.Of(anchor)
                    : throw new NotSupportedException("This chart is not anchored to a block of cells (it uses a fixed position), so it has no cell position.");
            }
        }

        NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor Anchor()
        {
            var rel = _drawing.RelationParts.FirstOrDefault(r => ReferenceEquals(r.DocumentPart, _chart));
            return rel == null ? null : _drawing.GetCTDrawing().CellAnchors.OfType<NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor>().FirstOrDefault(a =>
                a.graphicFrame?.graphic?.graphicData?.Any != null && a.graphicFrame.graphic.graphicData.Any.Any(xml =>
                    System.Text.RegularExpressions.Regex.Match(xml, "r:id=\"([^\"]+)\"").Groups[1].Value == rel.Relationship.Id));
        }

        /// <summary>
        /// NPOI's own RemoveChart needs the graphic frame it remembers from CreateChart, which a chart read from a file does not have.
        /// So: drop the anchor, then the relationship (a protected member of NPOI 2.7.6, which this library pins).
        /// </summary>
        internal void Remove()
        {
            var anchor = Anchor();
            if (anchor != null) _drawing.GetCTDrawing().CellAnchors.Remove(anchor);
            var drop = typeof(NPOI.POIXMLDocumentPart).GetMethod("RemoveRelation", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public,
                null, new[] { typeof(NPOI.POIXMLDocumentPart), typeof(bool) }, null)
                ?? throw new NotSupportedException("This version of NPOI cannot remove a chart that was read from a file.");
            drop.Invoke(_drawing, new object[] { _chart, true });
        }
        public IChartSeries AddSeries(string xRange, string yRange) => throw Frozen();
        public IChartSeries AddSeries(string values) => throw Frozen();
        public void Plot() => throw Frozen();
        public void SetTitle(string title) => _chart.SetTitle(title);

        public void SetLegendPosition(LegendPosition position)
        {
            var ch = _chart.GetCTChart();
            if (position == LegendPosition.None) { ch.legend = null; return; }
            ch.legend = ch.legend ?? new CT_Legend { overlay = new CT_Boolean { val = 0 } };
            ch.legend.legendPos = new CT_LegendPos { val = position == LegendPosition.Bottom ? ST_LegendPos.b : position == LegendPosition.Left ? ST_LegendPos.l : position == LegendPosition.Top ? ST_LegendPos.t : position == LegendPosition.TopRight ? ST_LegendPos.tr : ST_LegendPos.r };
            ch.legend.legendPosSpecified = true;
        }

        static NotSupportedException Frozen() => new NotSupportedException("Series of a chart loaded from a file cannot be changed; create a new chart instead.");
    }
}
