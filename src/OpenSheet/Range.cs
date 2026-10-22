using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Cysharp.Text;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal;
using OpenSheet.Styles;
using XL = NPOI.SS.UserModel;

namespace OpenSheet
{
    /// <summary>A rectangular block of cells on one worksheet. Indexes are zero-based.</summary>
    public class Range : IEnumerable<Cell>, IStylable
    {
        readonly WorkSheet _sheet;
        RangeAddress _addr;

        internal Range(WorkSheet sheet, RangeAddress address) { _sheet = sheet; _addr = address; }
        internal Range() { } // WorkSheet: its address is computed

        internal virtual WorkSheet Sheet => _sheet;

        public virtual RangeAddress RangeAddress => _addr;
        public virtual string RangeAddressAsString => RangeAddress.Location;
        public WorkBook WorkBook => Sheet.Book;
        public IStyle Style => new RangeStyle(this);

        public int Count => (int)Math.Min(int.MaxValue, (long)RowCount * ColumnCount);
        public int RowCount => RangeAddress.RowsCount;
        public int ColumnCount => RangeAddress.ColumnsCount;
        public bool IsEmpty => !Existing().Any(c => c.CellType != XL.CellType.Blank);

        public RangeRow[] Rows
        {
            get
            {
                var a = RangeAddress;
                return Enumerable.Range(a.FirstRow, a.RowsCount).Select(r => new RangeRow(Sheet, new RangeAddress(r, r, a.FirstColumn, a.LastColumn))).ToArray();
            }
        }

        public RangeColumn[] Columns
        {
            get
            {
                var a = RangeAddress;
                return Enumerable.Range(a.FirstColumn, a.ColumnsCount).Select(c => new RangeColumn(Sheet, new RangeAddress(a.FirstRow, a.LastRow, c, c))).ToArray();
            }
        }

        [Obsolete("Use Rows instead.")]
        public RangeRow[] AllRowsInRange => Rows;

        [Obsolete("Use Columns instead.")]
        public RangeColumn[] AllColumnsInRange => Columns;

        /// <summary>The sheet row with this zero-based index, across the columns of this range.</summary>
        protected RangeRow Row(int index)
        {
            var a = RangeAddress;
            return new RangeRow(Sheet, new RangeAddress(index, index, a.FirstColumn, a.LastColumn));
        }

        /// <summary>The sheet column with this zero-based index, across the rows of this range.</summary>
        protected RangeColumn Column(int index)
        {
            var a = RangeAddress;
            return new RangeColumn(Sheet, new RangeAddress(a.FirstRow, a.LastRow, index, index));
        }

        // ---- iteration helpers ----

        internal IEnumerable<Cell> AllCells()
        {
            var a = RangeAddress;
            for (var r = a.FirstRow; r <= a.LastRow; r++)
                for (var c = a.FirstColumn; c <= a.LastColumn; c++)
                    yield return new Cell(Sheet, r, c);
        }

        /// <summary>Cells that physically exist inside the range (skips the empty bulk of sparse sheets).</summary>
        internal IEnumerable<XL.ICell> Existing()
        {
            var a = RangeAddress;
            var sh = Sheet.Native;
            if (sh.PhysicalNumberOfRows == 0 || a.RowsCount <= 0) yield break;
            for (var r = Math.Max(a.FirstRow, sh.FirstRowNum); r <= Math.Min(a.LastRow, sh.LastRowNum); r++)
            {
                var row = sh.GetRow(r);
                if (row == null || row.FirstCellNum < 0) continue;
                for (var c = Math.Max(a.FirstColumn, row.FirstCellNum); c <= Math.Min(a.LastColumn, row.LastCellNum - 1); c++)
                {
                    var cell = row.GetCell(c);
                    if (cell != null) yield return cell;
                }
            }
        }

        internal IEnumerable<XL.ICell> EnsureCells() => AllCells().Select(c => c.Ensure());

        internal XL.ICell PeekFirst() => Sheet.Native.GetRow(RangeAddress.FirstRow)?.GetCell(RangeAddress.FirstColumn);

        Cell First => new Cell(Sheet, RangeAddress.FirstRow, RangeAddress.FirstColumn);

        IEnumerable<double> Numbers()
        {
            foreach (var c in Existing())
                if (CellValues.TryNumber(Sheet.Book, c, out var number)) yield return number;
        }

        // ---- values ----

        /// <summary>
        /// One cell: that cell's value. Several cells: the sum of the numeric ones.
        /// Setting writes every cell; null, "" or whitespace removes the cells.
        /// </summary>
        public object Value
        {
            get => Count == 1 ? First.Value : (object)Sum();
            set
            {
                if (CellValues.IsClearing(value)) { Clear(); return; }
                foreach (var c in AllCells()) c.Value = value;
            }
        }

        void Clear()
        {
            foreach (var c in Existing().ToList()) c.Row.RemoveCell(c);
            Sheet.Book.Touch();
        }

        public virtual void ClearContents()
        {
            foreach (var c in Existing()) c.SetBlank();
            Sheet.Book.Touch();
        }

        /// <summary>Display text of every cell, tab between columns and a new line between rows.</summary>
        public virtual string StringValue
        {
            get
            {
                var a = RangeAddress;
                var sheet = Sheet;
                var sb = ZString.CreateStringBuilder();
                try
                {
                    for (var r = a.FirstRow; r <= a.LastRow; r++)
                    {
                        if (r > a.FirstRow) sb.Append(Environment.NewLine);
                        var row = sheet.Native.GetRow(r);
                        for (var c = a.FirstColumn; c <= a.LastColumn; c++)
                        {
                            if (c > a.FirstColumn) sb.Append('\t');
                            sb.Append(Display.Formatted(sheet.Book, row?.GetCell(c)));
                        }
                    }
                    return sb.ToString();
                }
                finally { sb.Dispose(); }
            }
            set => Value = value;
        }

        /// <summary>The latest DateTime in the range, or null when there is none.</summary>
        public DateTime? DateTimeValue
        {
            get
            {
                DateTime? max = null;
                foreach (var c in Existing())
                    if (CellValues.Read(Sheet.Book, c, out var k) is DateTime d && k == Kind.Date && (max == null || d > max)) max = d;
                return max;
            }
            set => Value = value;
        }

        /// <summary>True only when every cell is true (numbers above zero count as true; empty cells do not).</summary>
        public virtual bool BoolValue
        {
            get { foreach (var c in AllCells()) if (!c.BoolValue) return false; return true; }
            set => Value = value;
        }

        public virtual double DoubleValue { get => (double)Sum(); set => Value = value; }
        public virtual decimal DecimalValue { get => Sum(); set => Value = value; }
        public virtual float FloatValue { get => (float)Sum(); set => Value = value; }
        public virtual int IntValue { get => Conv.ToInt((double)Sum()); set => Value = value; }
        public virtual long LongValue { get => Conv.ToLong((double)Sum()); set => Value = value; }
        public virtual long Int64Value { get => LongValue; set => LongValue = value; }
        public virtual int Int32Value { get => IntValue; set => IntValue = value; }

        public TType GetValue<TType>() => Count == 1 ? First.GetValue<TType>() : Conv.To<TType>(Sum());

        public bool TryGetValue<TType>(out TType value)
        {
            try { value = GetValue<TType>(); return true; }
            catch (Exception) { value = default(TType); return false; }
        }

        public override string ToString() => StringValue;

        /// <summary>Number format shared by all cells, or null when they differ. Setting formats every cell.</summary>
        public string FormatString
        {
            get
            {
                string common = null; var first = true;
                foreach (var c in AllCells())
                {
                    var f = c.FormatString;
                    if (first) { common = f; first = false; } else if (f != common) return null;
                }
                return common;
            }
            set
            {
                // what Cell.FormatString does, with the format looked up once for the whole range instead of once per cell
                var book = Sheet.Book;
                var fmt = value ?? "General";
                var idx = book.Native.CreateDataFormat().GetFormat(fmt);
                var key = "numfmt:" + fmt;
                Action<XL.ICellStyle> setFormat = s => s.DataFormat = idx;
                foreach (var c in AllCells()) book.Styles.Apply(c.Ensure(), key, setFormat);
                book.Touch();
            }
        }

        /// <summary>Formula shared by all cells, or null when they differ. Setting writes the same text to every cell.</summary>
        public string Formula
        {
            get
            {
                string common = null; var first = true;
                foreach (var c in AllCells())
                {
                    var f = c.Formula;
                    if (first) { common = f; first = false; } else if (f != common) return null;
                }
                return common;
            }
            set { foreach (var c in AllCells()) c.Formula = value; }
        }

        public decimal Sum() { decimal s = 0; foreach (var d in Numbers()) s += (decimal)d; return s; }
        public decimal Max() { var any = false; double m = 0; foreach (var d in Numbers()) { if (!any || d > m) m = d; any = true; } return any ? (decimal)m : 0; }
        public decimal Min() { var any = false; double m = 0; foreach (var d in Numbers()) { if (!any || d < m) m = d; any = true; } return any ? (decimal)m : 0; }
        public decimal Avg() { decimal s = 0; var n = 0; foreach (var d in Numbers()) { s += (decimal)d; n++; } return n == 0 ? 0 : s / n; }

        public virtual void SetCellValue(int rowIndex, int columnIndex, object value) => new Cell(Sheet, rowIndex, columnIndex).Value = value;

        public IEnumerator<Cell> GetEnumerator() => AllCells().GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public Cell[] ToArray() => AllCells().ToArray();

        // ---- combining ranges ----

        public static Range operator +(Range a, Range b) => new Range(a.Sheet, a.RangeAddress.Extend(b.RangeAddress));

        public static Range operator -(Range a, Range b)
        {
            var copy = new Range(a.Sheet, a.RangeAddress);
            return copy.SubtractRange(b);
        }

        /// <summary>Grows this range to cover both ranges (A1:B2 + C3:D4 = A1:D4).</summary>
        public virtual Range AddRange(Range range) => AddRange(range.RangeAddress);

        public virtual Range AddRange(RangeAddress address)
        {
            _addr = RangeAddress.Extend(address);
            return this;
        }

        public virtual Range SubtractRange(Range range) => SubtractRange(range.RangeAddress);

        /// <summary>
        /// Removes <paramref name="address"/> from this range. A slice that spans the full width or height from an edge
        /// shrinks the range; any other overlap is emptied but the range keeps its shape; full overlap leaves it with no cells.
        /// </summary>
        public virtual Range SubtractRange(RangeAddress address)
        {
            var a = RangeAddress;
            var x = a.GetIntersection(address);
            if (x == null) return this;
            var fullWidth = x.FirstColumn == a.FirstColumn && x.LastColumn == a.LastColumn;
            var fullHeight = x.FirstRow == a.FirstRow && x.LastRow == a.LastRow;
            if (fullWidth && fullHeight) { _addr = RangeAddress.None(a.FirstRow, a.FirstColumn); }
            else if (fullWidth && x.FirstRow == a.FirstRow) _addr = new RangeAddress(x.LastRow + 1, a.LastRow, a.FirstColumn, a.LastColumn);
            else if (fullWidth && x.LastRow == a.LastRow) _addr = new RangeAddress(a.FirstRow, x.FirstRow - 1, a.FirstColumn, a.LastColumn);
            else if (fullHeight && x.FirstColumn == a.FirstColumn) _addr = new RangeAddress(a.FirstRow, a.LastRow, x.LastColumn + 1, a.LastColumn);
            else if (fullHeight && x.LastColumn == a.LastColumn) _addr = new RangeAddress(a.FirstRow, a.LastRow, a.FirstColumn, x.FirstColumn - 1);
            else new Range(Sheet, x).ClearContents();
            return this;
        }

        /// <summary>A copy of this range without empty rows and columns on its borders.</summary>
        public Range Trim()
        {
            var a = RangeAddress;
            int r1 = int.MaxValue, r2 = -1, c1 = int.MaxValue, c2 = -1;
            foreach (var c in Existing())
            {
                if (c.CellType == XL.CellType.Blank) continue;
                r1 = Math.Min(r1, c.RowIndex); r2 = Math.Max(r2, c.RowIndex);
                c1 = Math.Min(c1, c.ColumnIndex); c2 = Math.Max(c2, c.ColumnIndex);
            }
            return new Range(Sheet, r2 < 0 ? RangeAddress.None(a.FirstRow, a.FirstColumn) : new RangeAddress(r1, r2, c1, c2));
        }

        // ---- rows and columns ----

        /// <summary>Empties the row inside this range; when nothing is left on that sheet row, the rows below move up.</summary>
        public virtual void RemoveRow(int rowNumber)
        {
            var a = RangeAddress;
            // a worksheet is every row of the sheet, not only the used ones
            if (this is WorkSheet ? rowNumber < 0 : rowNumber < a.FirstRow || rowNumber > a.LastRow) throw new ArgumentOutOfRangeException(nameof(rowNumber), "The row is outside this range.");
            foreach (var c in new Range(Sheet, new RangeAddress(rowNumber, rowNumber, a.FirstColumn, a.LastColumn)).Existing().ToList()) c.Row.RemoveCell(c);
            var sh = Sheet.Native;
            var row = sh.GetRow(rowNumber);
            if (row == null || row.PhysicalNumberOfCells == 0)
            {
                if (row != null) sh.RemoveRow(row);
                if (rowNumber < sh.LastRowNum) sh.ShiftRows(rowNumber + 1, sh.LastRowNum, -1);
                FeatureRanges.Adjust(Sheet, true, rowNumber, -1, includeConditionalFormats: false); // NPOI follows the conditional formats itself
                if (!(this is WorkSheet)) _addr = a.RowsCount == 1 ? RangeAddress.None(a.FirstRow, a.FirstColumn) : new RangeAddress(a.FirstRow, a.LastRow - 1, a.FirstColumn, a.LastColumn);
            }
            Sheet.Book.Touch();
        }

        public virtual void RemoveRow(RangeRow row) => RemoveRow(row.RowNumber);

        public virtual void RemoveColumn(string columnName) => RemoveColumn(A1.Column(columnName));
        public virtual void RemoveColumn(RangeColumn column) => RemoveColumn(column.ColumnNumber);

        /// <summary>Empties the column inside this range; when nothing is left in that sheet column, the columns to its right move left.</summary>
        public virtual void RemoveColumn(int columnNumber)
        {
            var a = RangeAddress;
            if (this is WorkSheet ? columnNumber < 0 : columnNumber < a.FirstColumn || columnNumber > a.LastColumn) throw new ArgumentOutOfRangeException(nameof(columnNumber), "The column is outside this range.");
            foreach (var c in new Range(Sheet, new RangeAddress(a.FirstRow, a.LastRow, columnNumber, columnNumber)).Existing().ToList()) c.Row.RemoveCell(c);
            if (!Sheet.ColumnHasCells(columnNumber))
            {
                Sheet.ShiftColumns(columnNumber + 1, -1);
                if (!(this is WorkSheet)) _addr = a.ColumnsCount == 1 ? RangeAddress.None(a.FirstRow, a.FirstColumn) : new RangeAddress(a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn - 1);
            }
            Sheet.Book.Touch();
        }

        public virtual RangeRow InsertRow(int rowIndex) => InsertRows(rowIndex, 1)[0];

        public virtual List<RangeRow> InsertRows(int rowIndex, int rowsCount = 1)
        {
            if (rowsCount < 1) throw new ArgumentOutOfRangeException(nameof(rowsCount));
            var sh = Sheet.Native;
            if (rowIndex <= sh.LastRowNum) sh.ShiftRows(rowIndex, sh.LastRowNum, rowsCount);
            FeatureRanges.Adjust(Sheet, true, rowIndex, rowsCount, includeConditionalFormats: false);
            var a = RangeAddress;
            var list = new List<RangeRow>();
            for (var i = 0; i < rowsCount; i++)
            {
                if (sh.GetRow(rowIndex + i) == null) sh.CreateRow(rowIndex + i);
                list.Add(new RangeRow(Sheet, new RangeAddress(rowIndex + i, rowIndex + i, a.FirstColumn, a.LastColumn)));
            }
            if (!(this is WorkSheet) && rowIndex >= a.FirstRow && rowIndex <= a.LastRow) _addr = new RangeAddress(a.FirstRow, a.LastRow + rowsCount, a.FirstColumn, a.LastColumn);
            Sheet.Book.Touch();
            return list;
        }

        public virtual RangeColumn InsertColumn(int columnIndex) => InsertColumns(columnIndex, 1)[0];
        public virtual RangeColumn InsertColumn(string columnLetter) => InsertColumns(A1.Column(columnLetter), 1)[0];
        public virtual List<RangeColumn> InsertColumns(string columnLetter, int columnsCount = 1) => InsertColumns(A1.Column(columnLetter), columnsCount);

        public virtual List<RangeColumn> InsertColumns(int columnIndex, int columnsCount = 1)
        {
            if (columnsCount < 1) throw new ArgumentOutOfRangeException(nameof(columnsCount));
            Sheet.ShiftColumns(columnIndex, columnsCount);
            var a = RangeAddress;
            var list = Enumerable.Range(columnIndex, columnsCount).Select(c => new RangeColumn(Sheet, new RangeAddress(a.FirstRow, a.LastRow, c, c))).ToList();
            if (!(this is WorkSheet) && columnIndex >= a.FirstColumn && columnIndex <= a.LastColumn) _addr = new RangeAddress(a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn + columnsCount);
            Sheet.Book.Touch();
            return list;
        }

        // ---- sorting ----

        public Range SortAscending() => SortByColumn(RangeAddress.FirstColumn, SortOrder.Ascending);
        public Range SortDescending() => SortByColumn(RangeAddress.FirstColumn, SortOrder.Descending);
        public Range SortByColumn(string column, SortOrder sortOrder) => SortByColumn(A1.Column(column), sortOrder);

        /// <summary>Reorders the rows of this range by the values of a sheet column (zero-based). Empty cells sort last.</summary>
        public Range SortByColumn(int columnNumber, SortOrder sortOrder)
        {
            var a = RangeAddress;
            if (columnNumber < a.FirstColumn || columnNumber > a.LastColumn) throw new ArgumentOutOfRangeException(nameof(columnNumber), "The column is outside this range.");
            var rows = Enumerable.Range(a.FirstRow, a.RowsCount).Select(r => new SortRow(Sheet, r, a, columnNumber)).ToList();
            var ordered = (sortOrder == SortOrder.Ascending
                ? rows.OrderBy(r => r, SortRow.Comparer(false)) : rows.OrderBy(r => r, SortRow.Comparer(true))).ToList(); // OrderBy is stable
            for (var i = 0; i < ordered.Count; i++) ordered[i].WriteTo(Sheet, a.FirstRow + i, a);
            Sheet.Book.Touch();
            return this;
        }

        // ---- copy / replace ----

        public void Copy(WorkSheet targetWorkSheet, string pointToInsert)
        {
            if (targetWorkSheet == null) throw new ArgumentNullException(nameof(targetWorkSheet));
            A1.ParseCell(pointToInsert, out var r0, out var c0);
            var a = RangeAddress;
            for (var r = a.FirstRow; r <= a.LastRow; r++)
                for (var c = a.FirstColumn; c <= a.LastColumn; c++)
                    Sheet.CopyCellTo(r, c, targetWorkSheet, r0 + r - a.FirstRow, c0 + c - a.FirstColumn);
        }

        /// <summary>Replaces text in string cells and in the text form of numbers; formulas are left alone.</summary>
        public void Replace(string oldValue, string newValue)
        {
            if (string.IsNullOrEmpty(oldValue)) throw new ArgumentException("The text to find is empty.", nameof(oldValue));
            foreach (var c in Existing().ToList())
            {
                if (c.CellType == XL.CellType.String)
                {
                    var s = c.StringCellValue;
                    if (s.Contains(oldValue)) c.SetCellValue(s.Replace(oldValue, newValue ?? ""));
                }
                else if (c.CellType == XL.CellType.Numeric && !CellValues.IsDateFormatted(Sheet.Book, c))
                {
                    var s = CellValues.Raw(c.NumericCellValue);
                    if (!s.Contains(oldValue)) continue;
                    var t = s.Replace(oldValue, newValue ?? "");
                    if (double.TryParse(t, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var d)) c.SetCellValue(d);
                    else c.SetCellValue(t);
                }
            }
            Sheet.Book.Touch();
        }

        // ---- names, tables, data ----

        public void SaveAsNamedRange(string nameOfRange, bool globalName) => Sheet.AddNamedRange(nameOfRange, this, globalName);

        public void SaveAsNamedTable(string tableName, bool showFilter = true, TableStyle tableStyle = null, bool useFirstRowAsHeader = false) =>
            Sheet.AddNamedTable(tableName, this, showFilter, tableStyle, useFirstRowAsHeader);

        public DataTable ToDataTable(bool useFirstRowAsColumnNames = false) => TableConverter.ToDataTable(this, useFirstRowAsColumnNames, "Table");
    }
}
