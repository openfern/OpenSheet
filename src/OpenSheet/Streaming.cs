using System;
using System.Collections.Generic;
using System.Linq;
using OpenSheet.Internal;

namespace OpenSheet
{
    /// <summary>One cell read by forward-only streaming. A snapshot: it is not connected to any workbook and cannot write back.</summary>
    public sealed class StreamCell
    {
        public StreamCell(int columnIndex, StreamCellType type, object value, string text, bool isFormula = false)
        {
            if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            ColumnIndex = columnIndex; Type = type; Value = value; Text = text; IsFormula = isFormula;
        }

        /// <summary>Zero-based.</summary>
        public int ColumnIndex { get; }
        public string ColumnLetter => A1.Column(ColumnIndex);
        /// <summary>For a formula cell this is the type of its cached result.</summary>
        public StreamCellType Type { get; }
        /// <summary>double, DateTime, string, bool or error text; null when <see cref="Type"/> is Blank.</summary>
        public object Value { get; }
        public string Text { get; }
        public bool IsFormula { get; }
    }

    /// <summary>Which sheet to stream and whether its first row names the columns.</summary>
    public sealed class StreamReadOptions
    {
        public StreamReadOptions() { }
        /// <summary>Zero-based. Ignored when <see cref="SheetName"/> is set.</summary>
        public int SheetIndex { get; set; }
        public string SheetName { get; set; }
        /// <summary>When true the row at <see cref="HeaderRowIndex"/> names the columns and is not returned; earlier rows are skipped.</summary>
        public bool HasHeaderRow { get; set; }
        /// <summary>Zero-based sheet row holding the header.</summary>
        public int HeaderRowIndex { get; set; }
    }

    /// <summary>A row read by streaming. Sparse: only columns that held something are present.</summary>
    public sealed class StreamRow
    {
        readonly Dictionary<string, int> _names;
        readonly bool _sorted;

        public StreamRow(int rowNumber, IReadOnlyList<StreamCell> cells) : this(rowNumber, cells, null, null) { }

        internal StreamRow(int rowNumber, IReadOnlyList<StreamCell> cells, IReadOnlyList<string> header, Dictionary<string, int> names)
        {
            RowNumber = rowNumber;
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            Header = header;
            _names = names;
            _sorted = true;
            for (var i = 1; i < cells.Count; i++) if (cells[i].ColumnIndex <= cells[i - 1].ColumnIndex) { _sorted = false; break; }
        }

        /// <summary>Zero-based (Excel row 1 is 0).</summary>
        public int RowNumber { get; }
        public IReadOnlyList<StreamCell> Cells { get; }
        /// <summary>Header[i] is the name at column i (null where blank), or null when no header was configured.</summary>
        public IReadOnlyList<string> Header { get; }

        /// <summary>The cell at that zero-based column, or null for a gap.</summary>
        public StreamCell this[int columnIndex]
        {
            get
            {
                if (!_sorted) return Cells.FirstOrDefault(c => c.ColumnIndex == columnIndex);
                int lo = 0, hi = Cells.Count - 1;
                while (lo <= hi)
                {
                    var mid = (lo + hi) >> 1;
                    var c = Cells[mid].ColumnIndex;
                    if (c == columnIndex) return Cells[mid];
                    if (c < columnIndex) lo = mid + 1; else hi = mid - 1;
                }
                return null;
            }
        }

        /// <summary>The cell under the named header column, or null for a gap. Throws when the name is unknown or there is no header.</summary>
        public StreamCell this[string columnName]
        {
            get
            {
                if (!TryGetCell(columnName, out var cell)) throw new KeyNotFoundException($"The header has no column named '{columnName}'.");
                return cell;
            }
        }

        /// <summary>False when the name is not in the header (cell is null). Throws when no header was configured.</summary>
        public bool TryGetCell(string columnName, out StreamCell cell)
        {
            if (_names == null) throw new InvalidOperationException("No header row was configured; set StreamReadOptions.HasHeaderRow.");
            if (columnName == null) throw new ArgumentNullException(nameof(columnName));
            if (_names.TryGetValue(columnName, out var idx)) { cell = this[idx]; return true; }
            cell = null;
            return false;
        }
    }
}
