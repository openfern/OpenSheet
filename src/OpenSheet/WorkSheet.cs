using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal;
using OpenSheet.Internal.Formulas;
using OpenSheet.Layout;
using OpenSheet.Options;
using OpenSheet.Printing;
using OpenSheet.Styles;
using XL = NPOI.SS.UserModel;

namespace OpenSheet
{
    /// <summary>One worksheet. It is also a <see cref="Range"/> covering all cells in use.</summary>
    public partial class WorkSheet : Range
    {
        internal WorkSheet(WorkBook book, XL.ISheet native) { Book = book; Native = native; }

        internal WorkBook Book { get; }
        internal XL.ISheet Native { get; }
        internal override WorkSheet Sheet => this;

        /// <summary>The cells in use: from the first to the last row and column that hold anything. An empty sheet reports A1.</summary>
        public override RangeAddress RangeAddress
        {
            get
            {
                if (Native.PhysicalNumberOfRows == 0) return new RangeAddress(0, 0, 0, 0);
                int r1 = Native.FirstRowNum, r2 = Native.LastRowNum, c1 = int.MaxValue, c2 = -1;
                for (var r = r1; r <= r2; r++)
                {
                    var row = Native.GetRow(r);
                    if (row == null || row.FirstCellNum < 0) continue;
                    c1 = Math.Min(c1, row.FirstCellNum); c2 = Math.Max(c2, row.LastCellNum - 1);
                }
                return c2 < 0 ? new RangeAddress(r1, r2, 0, 0) : new RangeAddress(r1, r2, c1, c2);
            }
        }

        public override Range AddRange(RangeAddress address) => throw new NotSupportedException("A worksheet always covers its used cells.");
        public override Range SubtractRange(RangeAddress address) => throw new NotSupportedException("A worksheet always covers its used cells.");

        // ---- identity ----

        public int Index => Book.Native.GetSheetIndex(Native);

        public string Name
        {
            get => Native.SheetName;
            set { Book.Native.SetSheetName(Index, value); Book.Touch(); }
        }

        public HiddenState Hidden
        {
            get => (HiddenState)(int)Book.Native.GetSheetVisibility(Index);
            set => Book.Native.SetSheetVisibility(Index, (XL.SheetVisibility)(int)value);
        }

        public bool IsRightToLeft { get => Native.IsRightToLeft; set => Native.IsRightToLeft = value; }

        /// <summary>Rows that hold at least one cell.</summary>
        public int PhysicalRowCount => Native.PhysicalNumberOfRows;

        /// <summary>Columns that hold a cell or carry a setting of their own (hidden, a width other than the default, a style).</summary>
        public int PhysicalColumnCount
        {
            get
            {
                var cols = new HashSet<int>();
                foreach (var c in FilledNative()) cols.Add(c.ColumnIndex);
                foreach (var c in ConfiguredColumns()) cols.Add(c);
                return cols.Count;
            }
        }

        // Columns that have their own definition. A definition spanning a large block is the sheet-wide default, not a setting.
        IEnumerable<int> ConfiguredColumns()
        {
            if (Native is XSSFSheet x)
            {
                var groups = x.GetCTWorksheet().cols;
                if (groups == null) yield break;
                foreach (var group in groups)
                    foreach (var col in group.col)
                    {
                        if (!(col.hidden || col.customWidth || col.style > 0) || col.max - col.min >= 256) continue;
                        for (var c = (int)col.min - 1; c < col.max; c++) yield return c;
                    }
                yield break;
            }
            var defaultWidth = Native.DefaultColumnWidth * 256;
            for (var c = 0; c <= References.XlsMaxColumn; c++)
                if (Native.IsColumnHidden(c) || Native.GetColumnWidth(c) != defaultWidth || Native.GetColumnStyle(c) != null) yield return c;
        }

        // ---- cells and ranges ----

        public Range this[string rangeAddress] => GetRange(rangeAddress);

        /// <summary>
        /// A range such as "A1:C9". A whole column ("A:A", "B:D") or whole row ("1:1", "2:5") covers the rows or columns in use
        /// instead of the million rows a sheet could hold.
        /// </summary>
        public Range GetRange(string range)
        {
            var text = (range ?? "").Trim();
            var bang = text.LastIndexOf('!');
            if (bang >= 0) text = text.Substring(bang + 1);
            // the used range walks every row, so only whole-column and whole-row addresses pay for it
            if (A1.TryWhole(text, true, out var first, out var second))
            {
                var used = RangeAddress;
                int c1 = A1.Column(first), c2 = A1.Column(second);
                return new Range(this, new RangeAddress(used.FirstRow, used.LastRow, Math.Min(c1, c2), Math.Max(c1, c2)));
            }
            if (A1.TryWhole(text, false, out first, out second))
            {
                var used = RangeAddress;
                int r1 = int.Parse(first) - 1, r2 = int.Parse(second) - 1;
                if (r1 < 0 || r2 < 0 || r1 > A1.MaxRow || r2 > A1.MaxRow) throw new ArgumentException($"Invalid range address '{range}'.", nameof(range));
                return new Range(this, new RangeAddress(Math.Min(r1, r2), Math.Max(r1, r2), used.FirstColumn, used.LastColumn));
            }
            return new Range(this, A1.Parse(range));
        }

        public Cell GetCellAt(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || rowIndex > A1.MaxRow) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            if (columnIndex < 0 || columnIndex > A1.MaxColumn) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            return new Cell(this, rowIndex, columnIndex);
        }

        /// <summary>The row across the used columns.</summary>
        public RangeRow GetRow(int rowIndex)
        {
            var a = RangeAddress;
            return new RangeRow(this, new RangeAddress(rowIndex, rowIndex, a.FirstColumn, a.LastColumn));
        }

        public RangeColumn GetColumn(string columnName) => GetColumn(A1.Column(columnName));

        /// <summary>The column down the used rows.</summary>
        public RangeColumn GetColumn(int columnIndex)
        {
            var a = RangeAddress;
            return new RangeColumn(this, new RangeAddress(a.FirstRow, a.LastRow, columnIndex, columnIndex));
        }

        public Cell[][] ToMultiDimensionalArray()
        {
            var a = RangeAddress;
            return Enumerable.Range(a.FirstRow, a.RowsCount)
                .Select(r => Enumerable.Range(a.FirstColumn, a.ColumnsCount).Select(c => new Cell(this, r, c)).ToArray()).ToArray();
        }

        IEnumerable<XL.ICell> FilledNative()
        {
            foreach (var row in Native)
                if (row != null)
                    foreach (var c in row) yield return c;
        }

        /// <summary>Cells that physically exist (holding a value or only formatting).</summary>
        public ICollection<Cell> FilledCells => FilledNative().Select(c => new Cell(this, c.RowIndex, c.ColumnIndex)).ToList();

        // The rows of a sheet and the cells of a row come in ascending order, so the first and last cell are found without sorting all of them.
        public Cell FirstFilledCell
        {
            get
            {
                foreach (var c in FilledNative()) return new Cell(this, c.RowIndex, c.ColumnIndex);
                return null;
            }
        }

        public Cell LastFilledCell
        {
            get
            {
                for (var r = Native.LastRowNum; r >= Native.FirstRowNum && r >= 0; r--)
                {
                    var row = Native.GetRow(r);
                    if (row == null || row.LastCellNum <= 0) continue;
                    var c = row.GetCell(row.LastCellNum - 1);
                    if (c != null) return new Cell(this, c.RowIndex, c.ColumnIndex);
                }
                return null;
            }
        }

        // ---- merge, autofilter, panes ----

        public List<Range> GetMergedRegions() => Native.MergedRegions.Select(m => new Range(this, RangeAddress.From(m))).ToList();

        public int Merge(string range)
        {
            var i = Native.AddMergedRegion(A1.Parse(range).ToNpoi());
            Book.Touch();
            return i;
        }

        public void Unmerge(int index) => Native.RemoveMergedRegion(index);

        public void Unmerge(string range)
        {
            var a = A1.Parse(range);
            for (var i = 0; i < Native.NumMergedRegions; i++)
                if (RangeAddress.From(Native.GetMergedRegion(i)).IsEqual(a)) { Native.RemoveMergedRegion(i); return; }
        }

        public void SetAutoFilter(string range) => SetAutoFilter(A1.Parse(range));

        public void SetAutoFilter(int firstRow, int lastRow, int firstColumn, int lastColumn) =>
            SetAutoFilter(new RangeAddress(firstRow, lastRow, firstColumn, lastColumn));

        void SetAutoFilter(RangeAddress a)
        {
            // Excel refuses a sheet filter over a table (the table has its own) and asks to repair the file.
            var table = TablesOverlapping(a).FirstOrDefault();
            if (table != null) throw new InvalidOperationException($"{a} overlaps the table '{table.Name}', which has its own filter.");
            Native.SetAutoFilter(a.ToNpoi());
        }

        IEnumerable<XSSFTable> TablesOverlapping(RangeAddress a) =>
            Native is XSSFSheet x
                ? x.GetTables().Where(t => a.GetIntersection(new RangeAddress(t.StartRowIndex, t.EndRowIndex, t.StartColIndex, t.EndColIndex)) != null)
                : Enumerable.Empty<XSSFTable>();

        RangeAddress AutoFilterRange()
        {
            var af = (Native as XSSFSheet)?.GetCTWorksheet().autoFilter;
            return af?.@ref == null ? null : A1.Parse(af.@ref);
        }

        public void UnsetAutoFilter()
        {
            if (!(Native is XSSFSheet x)) throw new NotSupportedException("Removing an auto filter is supported for the XLSX format only.");
            var ct = x.GetCTWorksheet();
            if (ct.autoFilter != null) ct.autoFilter = null;
            var name = Book.Native.GetNames("_xlnm._FilterDatabase").FirstOrDefault(n => n.SheetIndex == Index);
            if (name != null) Book.Native.RemoveName(name);
        }

        public void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) => Native.CreateFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
        public void CreateFreezePane(int colSplit, int rowSplit) => Native.CreateFreezePane(colSplit, rowSplit);

        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane) =>
            Native.CreateSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, (XL.PanePosition)(int)activePane);

        public void RemovePane() => Native.CreateFreezePane(0, 0);

        public void SetZoom(int scale) => Native.SetZoom(scale);

        public void SetRowBreak(int row) => Native.SetRowBreak(row);
        public void RemoveRowBreak(int row) => Native.RemoveRowBreak(row);
        public void SetColumnBreak(int column) => Native.SetColumnBreak(column);
        public void RemoveColumnBreak(int column) => Native.RemoveColumnBreak(column);

        // ---- display / print ----

        public bool DisplayGridlines { get => Native.DisplayGridlines; set => Native.DisplayGridlines = value; }
        public bool IsPrintGridlines { get => Native.IsPrintGridlines; set => Native.IsPrintGridlines = value; }
        public bool FitToPage { get => Native.FitToPage; set => Native.FitToPage = value; }
        public bool HorizontallyCenter { get => Native.HorizontallyCenter; set => Native.HorizontallyCenter = value; }
        public bool VerticallyCenter { get => Native.VerticallyCenter; set => Native.VerticallyCenter = value; }
        public IPrintSetup PrintSetup => new PrintSetupAdapter(Native.PrintSetup);
        public IHeader Header => new HeaderFooterAdapter(Native.Header);
        public IFooter Footer => new HeaderFooterAdapter(Native.Footer);

        // NPOI prefixes the sheet name itself, so only the area is passed.
        public void SetPrintArea(string range) => Book.Native.SetPrintArea(Index, Absolute(A1.Parse(range)));

        public void SetRepeatingRows(int startRow, int endRow) => Native.RepeatingRows = new CellRangeAddress(startRow, endRow, -1, -1);
        public void SetRepeatingColumns(int startColumn, int endColumn) => Native.RepeatingColumns = new CellRangeAddress(-1, -1, startColumn, endColumn);
        public void SetRepeatingColumns(string startColumn, string endColumn) => SetRepeatingColumns(A1.Column(startColumn), A1.Column(endColumn));
        public void RemoveRepeatingRows() => Native.RepeatingRows = null;
        public void RemoveRepeatingColumns() => Native.RepeatingColumns = null;

        /// <summary>Tab color as RGB bytes, or null when none is set.</summary>
        public byte[] TabColor
        {
            get
            {
                if (Native is XSSFSheet x) { var c = x.TabColor; var rgb = c?.RGB; return rgb == null ? null : rgb.Length == 4 ? new[] { rgb[1], rgb[2], rgb[3] } : rgb; }
                if (Native is HSSFSheet h && !h.IsAutoTabColor) return Book.Hssf.GetCustomPalette().GetColor(h.TabColorIndex)?.RGB;
                return null;
            }
        }

        /// <summary>Works for XLSX; for XLS use <see cref="SetWorksheetTabColorIndex"/>.</summary>
        public void SetWorksheetTabColor(byte[] rgb)
        {
            if (rgb == null || rgb.Length != 3) throw new ArgumentException("Pass exactly three bytes: red, green, blue.", nameof(rgb));
            if (Native is XSSFSheet x) x.TabColor = Colors.Xssf(rgb);
            else throw new NotSupportedException("RGB tab colors need the XLSX format; use SetWorksheetTabColorIndex for XLS.");
        }

        public void SetWorksheetTabColorIndex(short colorIndex)
        {
            if (Native is HSSFSheet h) h.TabColorIndex = colorIndex;
            else if (Native is XSSFSheet x) x.TabColor = new XSSFColor(XL.IndexedColors.ValueOf(colorIndex), null);
        }

        // ---- grouping / protection / size ----

        public void GroupRows(int startRow, int endRow) => Native.GroupRow(startRow, endRow);
        public void GroupColumns(int startColumn, int endColumn) => Native.GroupColumn(startColumn, endColumn);
        public void GroupColumns(string startColumn, string endColumn) => GroupColumns(A1.Column(startColumn), A1.Column(endColumn));
        public void UngroupRows(int startRow, int endRow) => Native.UngroupRow(startRow, endRow);
        public void UngroupColumns(int startColumn, int endColumn) => Native.UngroupColumn(startColumn, endColumn);
        public void UngroupColumn(string startColumn, string endColumn) => UngroupColumns(A1.Column(startColumn), A1.Column(endColumn));

        public void ProtectSheet(string password) => Native.ProtectSheet(password);
        public void UnprotectSheet() => Native.ProtectSheet(null);

        public void AutoSizeColumn(int columnIndex) => Native.AutoSizeColumn(columnIndex);
        public void AutoSizeColumn(int columnIndex, bool useMergedCells) => Native.AutoSizeColumn(columnIndex, useMergedCells);
        public void AutoSizeRow(int rowIndex) => Native.AutoSizeRow(rowIndex);
        public void AutoSizeRow(int rowIndex, bool useMergedCells) => Native.AutoSizeRow(rowIndex, useMergedCells);

        // ---- copying ----

        public void CopyTo(WorkBook targetWorkbook, string name)
        {
            if (targetWorkbook == null) throw new ArgumentNullException(nameof(targetWorkbook));
            SheetCopier.Copy(this, targetWorkbook, name);
        }

        public WorkSheet CopySheet(string name)
        {
            SheetCopier.CopyNative(this, Book.Native, name);
            Book.Touch();
            return Book.Wrap(Book.Native.GetSheet(name));
        }

        internal void CopyCellTo(int sr, int sc, WorkSheet dst, int r, int c)
        {
            var src = Native.GetRow(sr)?.GetCell(sc);
            var drow = dst.Native.GetRow(r) ?? dst.Native.CreateRow(r);
            var old = drow.GetCell(c);
            if (old != null) drow.RemoveCell(old);
            if (src == null) return;
            var d = drow.CreateCell(c);
            switch (src.CellType)
            {
                case XL.CellType.Numeric: d.SetCellValue(src.NumericCellValue); break;
                case XL.CellType.String: d.SetCellValue(src.StringCellValue); break;
                case XL.CellType.Boolean: d.SetCellValue(src.BooleanCellValue); break;
                case XL.CellType.Error: d.SetCellErrorValue(src.ErrorCellValue); break;
                // a copied formula keeps pointing at the same neighbours: relative references move with it, $ ones stay
                case XL.CellType.Formula: d.SetCellFormula(References.Translate(src.CellFormula, r - sr, c - sc, dst.MaxRow, dst.MaxColumn)); break;
            }
            d.CellStyle = dst.Book.ImportStyle(src.CellStyle, Book);
            dst.Book.Touch();
        }

        internal int MaxRow => Book.Format == ExcelFileFormat.XLS ? References.XlsMaxRow : References.XlsxMaxRow;
        internal int MaxColumn => Book.Format == ExcelFileFormat.XLS ? References.XlsMaxColumn : References.XlsxMaxColumn;

        internal bool ColumnHasCells(int column) => FilledNative().Any(c => c.ColumnIndex == column);

        /// <summary>
        /// Moves every cell at or right of <paramref name="start"/> by <paramref name="n"/> columns (negative moves left), and fixes
        /// what points at them: formulas on every sheet, defined names, merged regions and the auto filter.
        /// </summary>
        internal void ShiftColumns(int start, int n)
        {
            if (n == 0) return;
            var last = -1;
            foreach (var row in Native) if (row != null) last = Math.Max(last, row.LastCellNum - 1);
            if (last >= start)
            {
                foreach (var row in Native.Cast<XL.IRow>().Where(r => r != null).ToList())
                {
                    var cols = n > 0 ? Enumerable.Range(start, last - start + 1).Reverse() : Enumerable.Range(start, last - start + 1);
                    foreach (var c in cols)
                    {
                        var cell = row.GetCell(c);
                        if (cell == null) continue;
                        var target = row.GetCell(c + n);
                        if (target != null) row.RemoveCell(target);
                        // NPOI's XSSFRow.MoveCell is not implemented; copy value, style, link and comment, then drop the original.
                        var moved = row.CreateCell(c + n);
                        SheetCopier.CopyValue(cell, moved);
                        moved.CellStyle = cell.CellStyle;
                        if (cell.Hyperlink != null)
                        {
                            var link = Book.Native.GetCreationHelper().CreateHyperlink(cell.Hyperlink.Type);
                            link.Address = cell.Hyperlink.Address;
                            moved.Hyperlink = link;
                        }
                        var comment = cell.CellComment;
                        row.RemoveCell(cell);
                        if (comment != null) moved.CellComment = comment;
                    }
                }
                var widths = Enumerable.Range(start, last - start + 1).Select(c => (c, w: Native.GetColumnWidth(c), h: Native.IsColumnHidden(c))).ToList();
                foreach (var (c, w, h) in n > 0 ? widths.AsEnumerable().Reverse() : widths)
                {
                    Native.SetColumnWidth(c + n, w);
                    Native.SetColumnHidden(c + n, h);
                }
            }
            AdjustReferences(n > 0 ? start : start + n, n);
            Book.Touch();
        }

        // Columns were inserted before `at` (delta > 0) or -delta of them were removed from `at`.
        void AdjustReferences(int at, int delta)
        {
            var maxRow = MaxRow; var maxColumn = MaxColumn;
            foreach (var sheet in Book.SheetList())
                foreach (var row in sheet.Native.Cast<XL.IRow>().Where(r => r != null))
                    foreach (var cell in row.Where(c => c.CellType == XL.CellType.Formula).ToList())
                    {
                        var before = cell.CellFormula;
                        var after = References.Adjust(before, sheet.Name, Name, false, at, delta, maxRow, maxColumn);
                        if (after != before) cell.SetCellFormula(after);
                    }

            foreach (var name in Book.Native.GetAllNames().Where(nm => !nm.IsFunctionName && !nm.IsDeleted && !string.IsNullOrEmpty(nm.RefersToFormula)).ToList())
            {
                var owner = name.SheetIndex >= 0 ? Book.Native.GetSheetName(name.SheetIndex) : null;
                var after = References.Adjust(name.RefersToFormula, owner, Name, false, at, delta, maxRow, maxColumn);
                if (after != name.RefersToFormula) name.RefersToFormula = after;
            }

            var regions = Native.MergedRegions.ToList();
            for (var i = Native.NumMergedRegions - 1; i >= 0; i--) Native.RemoveMergedRegion(i);
            foreach (var region in regions)
            {
                int first = region.FirstColumn, lastColumn = region.LastColumn;
                if (!References.AdjustBlock(ref first, ref lastColumn, at, delta, maxColumn) || (first == lastColumn && region.FirstRow == region.LastRow)) continue;
                Native.AddMergedRegion(new CellRangeAddress(region.FirstRow, region.LastRow, first, lastColumn));
            }

            FeatureRanges.Adjust(this, false, at, delta, includeConditionalFormats: true);

            var filter = AutoFilterRange();
            if (filter != null && Native is XSSFSheet x)
            {
                int first = filter.FirstColumn, lastColumn = filter.LastColumn;
                if (References.AdjustBlock(ref first, ref lastColumn, at, delta, maxColumn))
                    x.GetCTWorksheet().autoFilter.@ref = new RangeAddress(filter.FirstRow, filter.LastRow, first, lastColumn).Location;
            }
        }

        // ---- names and tables ----

        internal static string Quote(string sheetName) => "'" + sheetName.Replace("'", "''") + "'";

        internal static string Absolute(RangeAddress a) =>
            a.RowsCount == 1 && a.ColumnsCount == 1
                ? $"${A1.Column(a.FirstColumn)}${a.FirstRow + 1}"
                : $"${A1.Column(a.FirstColumn)}${a.FirstRow + 1}:${A1.Column(a.LastColumn)}${a.LastRow + 1}";

        public void AddNamedRange(string nameOfRange, Range range, bool globalName = true) =>
            AddNamedRange(nameOfRange, Quote(range.Sheet.Name) + "!" + Absolute(range.RangeAddress), globalName);

        public void AddNamedRange(string nameOfRange, string formula, bool globalName = true)
        {
            NameRules.Validate(nameOfRange, nameof(nameOfRange), "range");
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("The formula is empty.", nameof(formula));
            var n = Book.Native.CreateName();
            if (!globalName) n.SheetIndex = Index;
            n.NameName = nameOfRange;
            n.RefersToFormula = formula.Trim().TrimStart('=');
            Book.Touch();
        }

        XL.IName FindName(string name) =>
            Book.Native.GetNames(name).FirstOrDefault(n => n.SheetIndex == Index) ?? Book.Native.GetNames(name).FirstOrDefault(n => n.SheetIndex < 0);

        public string FindNamedRange(string nameOfRange) => FindName(nameOfRange)?.RefersToFormula;

        public void RemoveNamedRange(string nameOfRange)
        {
            var n = FindName(nameOfRange);
            if (n != null) Book.Native.RemoveName(n);
        }

        public ICollection<string> GetNamedRanges() =>
            Book.Native.GetAllNames()
                .Where(n => !n.IsFunctionName && !n.IsDeleted && !n.NameName.StartsWith("_xlnm", StringComparison.Ordinal) && (n.SheetIndex < 0 || n.SheetIndex == Index))
                .Select(n => n.NameName).ToList();

        XSSFSheet Xssf => Native as XSSFSheet ?? throw new NotSupportedException("Named tables need the XLSX format.");

        public void AddNamedTable(string tableName, Range range, bool showFilter = true, TableStyle tableStyle = null, bool useFirstRowAsHeader = false)
        {
            NameRules.Validate(tableName, nameof(tableName), "table");
            var a = range.RangeAddress;
            var sheet = Xssf;
            var clash = TablesOverlapping(a).FirstOrDefault();
            if (clash != null) throw new ArgumentException($"{a} overlaps the table '{clash.Name}'.", nameof(range));
            // Making a filtered range a table hands the filter to the table, as Excel does.
            if (AutoFilterRange()?.GetIntersection(a) != null) UnsetAutoFilter();
            var headers = Enumerable.Range(a.FirstColumn, a.ColumnsCount)
                .Select(c => useFirstRowAsHeader ? new Cell(this, a.FirstRow, c).StringValue : null).ToList();
            for (var i = 0; i < headers.Count; i++)
                if (string.IsNullOrWhiteSpace(headers[i]) || headers.Take(i).Contains(headers[i])) headers[i] = "Column" + (i + 1);

            var t = sheet.CreateTable();
            t.Name = tableName;
            t.DisplayName = tableName;
            t.CellReferences = new AreaReference(new CellReference(a.FirstRow, a.FirstColumn), new CellReference(a.LastRow, a.LastColumn));
            var ct = t.GetCTTable();
            ct.headerRowCount = useFirstRowAsHeader ? 1u : 0u;
            for (var i = 0; i < headers.Count; i++) t.CreateColumn(headers[i], i);
            if (showFilter && useFirstRowAsHeader) { ct.autoFilter = new NPOI.OpenXmlFormats.Spreadsheet.CT_AutoFilter { @ref = a.Location }; }
            if (tableStyle != null && tableStyle.Name != null)
            {
                ct.tableStyleInfo = new NPOI.OpenXmlFormats.Spreadsheet.CT_TableStyleInfo
                {
                    name = tableStyle.Name,
                    showFirstColumn = tableStyle.ShowFirstColumn ?? false,
                    showLastColumn = tableStyle.ShowLastColumn ?? false,
                    showRowStripes = tableStyle.ShowRowStripes ?? true,
                    showColumnStripes = tableStyle.ShowColumnStripes ?? false,
                };
            }
            Book.Touch();
        }

        XSSFTable FindTable(string name) =>
            Xssf.GetTables().FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException($"No table named '{name}' on sheet '{Name}'.", nameof(name));

        public ICollection<string> GetNamedTableNames() => Xssf.GetTables().Select(t => t.Name).ToList();

        public Range GetNamedTable(string tableName)
        {
            var t = FindTable(tableName);
            return new Range(this, new RangeAddress(t.StartRowIndex, t.EndRowIndex, t.StartColIndex, t.EndColIndex));
        }

        public DataTable GetDataTableFromNamedTable(string tableName, bool useFirstRowAsColumnNames = false) =>
            TableConverter.ToDataTable(GetNamedTable(tableName), useFirstRowAsColumnNames, tableName);

        // ---- export ----

        public WorkSheet SaveAs(string filename)
        {
            Book.SaveSheet(this, filename);
            return this;
        }

        public void SaveAsNewExcelFile(string filename, ExcelFileFormat excelFormat = ExcelFileFormat.XLSX) =>
            Book.SingleSheetBook(this, excelFormat).SaveAs(filename);

        public void SaveAsCsv(string filename, string delimiter = null) => File.WriteAllBytes(filename, ToCsvStream(delimiter).ToArray());
        public void SaveAsXml(string filename) => File.WriteAllBytes(filename, ToXmlStream().ToArray());
        public void SaveAsJson(string filename) => File.WriteAllBytes(filename, ToJsonStream().ToArray());

        public MemoryStream ToNewExcelStream(ExcelFileFormat excelFormat = ExcelFileFormat.XLSX) => Book.SingleSheetBook(this, excelFormat).ToStream(excelFormat);
        public MemoryStream ToCsvStream(string delimiter = null) => TextExport.Stream(TextExport.Csv(this, delimiter ?? ","));
        public MemoryStream ToTsvStream() => TextExport.Stream(TextExport.Csv(this, "\t"));
        public MemoryStream ToXmlStream() => TextExport.Stream(TextExport.Xml(new[] { this }));
        public MemoryStream ToJsonStream() => TextExport.Stream(TextExport.Json(this));
    }
}
