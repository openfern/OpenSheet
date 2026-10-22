using OpenSheet.Internal;

namespace OpenSheet
{
    /// <summary>One row of a range.</summary>
    public class RangeRow : Range
    {
        internal RangeRow(WorkSheet sheet, RangeAddress address) : base(sheet, address) { }

        /// <summary>Zero-based row index.</summary>
        public int RowNumber => RangeAddress.FirstRow;

        /// <summary>Row height in 1/20 of a point (Excel's native unit; 300 is the usual 15 pt row).</summary>
        public int Height
        {
            get => Sheet.Native.GetRow(RowNumber)?.Height ?? Sheet.Native.DefaultRowHeight;
            set
            {
                var row = Sheet.Native.GetRow(RowNumber) ?? Sheet.Native.CreateRow(RowNumber);
                row.Height = (short)System.Math.Max(0, System.Math.Min(short.MaxValue, value));
                Sheet.Book.Touch();
            }
        }

        public bool Hidden
        {
            get => Sheet.Native.GetRow(RowNumber)?.ZeroHeight ?? false;
            set { (Sheet.Native.GetRow(RowNumber) ?? Sheet.Native.CreateRow(RowNumber)).ZeroHeight = value; }
        }

        public void RemoveRow() => RemoveRow(RowNumber);
        public void AutoSizeRow() => Sheet.AutoSizeRow(RowNumber);
        public void AutoSizeRow(bool useMergedCells) => Sheet.AutoSizeRow(RowNumber, useMergedCells);
    }
}
