using System;
using System.Collections.Generic;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>A row captured before sorting, so rows can be rewritten in a new order without losing content or style.</summary>
    internal sealed class SortRow
    {
        sealed class Snap
        {
            public bool Exists;
            public XL.CellType Type;
            public double Num;
            public string Str;
            public bool Bool;
            public byte Err;
            public string Formula;
            public XL.ICellStyle Style;
        }

        readonly Snap[] _cells;
        readonly int _row;
        readonly object _key;
        readonly Kind _kind;

        public SortRow(WorkSheet sheet, int row, RangeAddress a, int keyColumn)
        {
            var r = sheet.Native.GetRow(row);
            _row = row;
            _cells = new Snap[a.ColumnsCount];
            for (var i = 0; i < _cells.Length; i++)
            {
                var c = r?.GetCell(a.FirstColumn + i);
                var s = new Snap();
                if (c != null)
                {
                    s.Exists = true; s.Style = c.CellStyle; s.Type = c.CellType;
                    switch (c.CellType)
                    {
                        case XL.CellType.Numeric: s.Num = c.NumericCellValue; break;
                        case XL.CellType.String: s.Str = c.StringCellValue; break;
                        case XL.CellType.Boolean: s.Bool = c.BooleanCellValue; break;
                        case XL.CellType.Error: s.Err = c.ErrorCellValue; break;
                        case XL.CellType.Formula: s.Formula = c.CellFormula; break;
                    }
                }
                _cells[i] = s;
            }
            _key = CellValues.Read(sheet.Book, r?.GetCell(keyColumn), out _kind);
        }

        static int Rank(Kind k)
        {
            switch (k)
            {
                case Kind.Number: case Kind.Date: return 0;
                case Kind.Text: return 1;
                case Kind.Bool: return 2;
                default: return 3; // errors
            }
        }

        /// <summary>Excel order: numbers, text, booleans, errors; empty cells always last.</summary>
        public static IComparer<SortRow> Comparer(bool descending) => new Cmp(descending);

        sealed class Cmp : IComparer<SortRow>
        {
            readonly bool _desc;
            public Cmp(bool desc) { _desc = desc; }

            public int Compare(SortRow x, SortRow y)
            {
                var ex = x._kind == Kind.Empty; var ey = y._kind == Kind.Empty;
                if (ex || ey) return ex == ey ? 0 : ex ? 1 : -1;
                var c = Rank(x._kind).CompareTo(Rank(y._kind));
                if (c == 0)
                {
                    switch (Rank(x._kind))
                    {
                        case 0: c = CellValues.ToDouble(x._key).CompareTo(CellValues.ToDouble(y._key)); break;
                        case 1: c = string.Compare((string)x._key, (string)y._key, StringComparison.InvariantCultureIgnoreCase); break;
                        case 2: c = ((bool)x._key).CompareTo((bool)y._key); break;
                        default: c = string.CompareOrdinal((string)x._key, (string)y._key); break;
                    }
                }
                return _desc ? -c : c;
            }
        }

        public void WriteTo(WorkSheet sheet, int row, RangeAddress a)
        {
            var r = sheet.Native.GetRow(row) ?? sheet.Native.CreateRow(row);
            for (var i = 0; i < _cells.Length; i++)
            {
                var col = a.FirstColumn + i;
                var s = _cells[i];
                var old = r.GetCell(col);
                if (old != null) r.RemoveCell(old);
                if (!s.Exists) continue;
                var c = r.CreateCell(col);
                switch (s.Type)
                {
                    case XL.CellType.Numeric: c.SetCellValue(s.Num); break;
                    case XL.CellType.String: c.SetCellValue(s.Str); break;
                    case XL.CellType.Boolean: c.SetCellValue(s.Bool); break;
                    case XL.CellType.Error: c.SetCellErrorValue(s.Err); break;
                    // the formula moves with its row: relative references follow it, as when Excel sorts
                    case XL.CellType.Formula: c.SetCellFormula(Formulas.References.Translate(s.Formula, row - _row, 0, sheet.MaxRow, sheet.MaxColumn)); break;
                }
                c.CellStyle = s.Style;
            }
        }
    }
}
