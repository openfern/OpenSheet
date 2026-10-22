using System;
using OpenSheet.Internal;

namespace OpenSheet
{
    /// <summary>A rectangular block of cells, described by zero-based first/last row and column indexes.</summary>
    public class RangeAddress
    {
        internal RangeAddress(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (firstRow > lastRow || firstColumn > lastColumn) throw new ArgumentException("First index must not exceed last index.");
            FirstRow = firstRow; LastRow = lastRow; FirstColumn = firstColumn; LastColumn = lastColumn;
        }

        RangeAddress(int row, int column)
        {
            FirstRow = row; LastRow = row - 1; FirstColumn = column; LastColumn = column - 1; // zero cells
        }

        /// <summary>A range that holds no cells, anchored where the original range started.</summary>
        internal static RangeAddress None(int row, int column) => new RangeAddress(row, column);

        public int FirstRow { get; }
        public int LastRow { get; }
        public int FirstColumn { get; }
        public int LastColumn { get; }
        public int ColumnsCount => LastColumn - FirstColumn + 1;
        public int RowsCount => LastRow - FirstRow + 1;

        /// <summary>Excel notation, e.g. "B2:D9". A single cell is written without the colon: "B2".</summary>
        public string Location
        {
            get
            {
                if (RowsCount <= 0 || ColumnsCount <= 0) return "";
                var a = A1.Cell(FirstRow, FirstColumn);
                return FirstRow == LastRow && FirstColumn == LastColumn ? a : a + ":" + A1.Cell(LastRow, LastColumn);
            }
        }

        /// <summary>Smallest address that contains both this address and <paramref name="rangeAddress"/>.</summary>
        public RangeAddress Extend(RangeAddress rangeAddress)
        {
            if (rangeAddress == null) throw new ArgumentNullException(nameof(rangeAddress));
            return new RangeAddress(Math.Min(FirstRow, rangeAddress.FirstRow), Math.Max(LastRow, rangeAddress.LastRow),
                Math.Min(FirstColumn, rangeAddress.FirstColumn), Math.Max(LastColumn, rangeAddress.LastColumn));
        }

        public bool Includes(int row, int col) => row >= FirstRow && row <= LastRow && col >= FirstColumn && col <= LastColumn;

        public bool IsEqual(RangeAddress range) =>
            range != null && FirstRow == range.FirstRow && LastRow == range.LastRow && FirstColumn == range.FirstColumn && LastColumn == range.LastColumn;

        /// <summary>The cells both addresses share, or null when they do not overlap.</summary>
        public RangeAddress GetIntersection(RangeAddress range)
        {
            if (range == null) return null;
            var r1 = Math.Max(FirstRow, range.FirstRow); var r2 = Math.Min(LastRow, range.LastRow);
            var c1 = Math.Max(FirstColumn, range.FirstColumn); var c2 = Math.Min(LastColumn, range.LastColumn);
            return r1 > r2 || c1 > c2 ? null : new RangeAddress(r1, r2, c1, c2);
        }

        public override string ToString() => Location;
        public override bool Equals(object obj) => obj is RangeAddress other && IsEqual(other);
        public override int GetHashCode() => (((FirstRow * 397) ^ LastRow) * 397 ^ FirstColumn) * 397 ^ LastColumn;

        internal NPOI.SS.Util.CellRangeAddress ToNpoi() => new NPOI.SS.Util.CellRangeAddress(FirstRow, LastRow, FirstColumn, LastColumn);
        internal static RangeAddress From(NPOI.SS.Util.CellRangeAddress a) => new RangeAddress(a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn);
    }
}
