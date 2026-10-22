using OpenSheet.Internal;

namespace OpenSheet
{
    /// <summary>One column of a range.</summary>
    public class RangeColumn : Range
    {
        internal RangeColumn(WorkSheet sheet, RangeAddress address) : base(sheet, address) { }

        /// <summary>Zero-based column index.</summary>
        public int ColumnNumber => RangeAddress.FirstColumn;
        public string ColumnLetter => A1.Column(ColumnNumber);

        /// <summary>Column width in 1/256 of a character (Excel's native unit; 2048 is 8 characters).</summary>
        public int Width
        {
            get => (int)Sheet.Native.GetColumnWidth(ColumnNumber);
            set { Sheet.Native.SetColumnWidth(ColumnNumber, value); Sheet.Book.Touch(); }
        }

        public bool Hidden
        {
            get => Sheet.Native.IsColumnHidden(ColumnNumber);
            set => Sheet.Native.SetColumnHidden(ColumnNumber, value);
        }

        public void RemoveColumn() => RemoveColumn(ColumnNumber);
        public void AutoSizeColumn() => Sheet.AutoSizeColumn(ColumnNumber);
        public void AutoSizeColumn(bool useMergedCells) => Sheet.AutoSizeColumn(ColumnNumber, useMergedCells);
    }
}
