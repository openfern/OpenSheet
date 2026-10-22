using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace OpenSheet.Internal
{
    /// <summary>Moves data between ranges and System.Data tables.</summary>
    internal static class TableConverter
    {
        /// <summary>
        /// Columns take the common CLR type of their values (double, DateTime, bool, string), or object when mixed.
        /// Header text becomes the column name; blank or repeated names become "Column{n}".
        /// </summary>
        public static DataTable ToDataTable(Range range, bool firstRowAsHeader, string tableName)
        {
            var a = range.RangeAddress;
            var table = new DataTable(tableName);
            if (a.RowsCount <= 0 || a.ColumnsCount <= 0) return table;
            var sheet = range.Sheet;
            var dataStart = a.FirstRow + (firstRowAsHeader ? 1 : 0);

            var names = new List<string>();
            var taken = new HashSet<string>();
            for (var c = 0; c < a.ColumnsCount; c++)
            {
                var n = firstRowAsHeader ? new Cell(sheet, a.FirstRow, a.FirstColumn + c).StringValue : null;
                if (string.IsNullOrWhiteSpace(n) || taken.Contains(n))
                {
                    n = "Column" + (c + 1);
                    while (taken.Contains(n)) n += "_"; // a real header may already be called Column3
                }
                taken.Add(n);
                names.Add(n);
            }

            var rows = new List<object[]>();
            var types = new Type[a.ColumnsCount];
            for (var r = dataStart; r <= a.LastRow; r++)
            {
                var values = new object[a.ColumnsCount];
                var row = sheet.Native.GetRow(r);
                for (var c = 0; c < values.Length; c++)
                {
                    var v = CellValues.Read(sheet.Book, row?.GetCell(a.FirstColumn + c), out _);
                    values[c] = v ?? DBNull.Value;
                    if (v == null) continue;
                    types[c] = types[c] == null || types[c] == v.GetType() ? v.GetType() : typeof(object);
                }
                rows.Add(values);
            }

            for (var c = 0; c < names.Count; c++) table.Columns.Add(names[c], types[c] ?? typeof(object));
            table.BeginLoadData();
            foreach (var v in rows) table.Rows.Add(v);
            table.EndLoadData();
            return table;
        }

        /// <summary>Column names in row 0, rows below; DBNull and null leave the cell empty.</summary>
        public static void Fill(WorkSheet sheet, DataTable table)
        {
            for (var c = 0; c < table.Columns.Count; c++) new Cell(sheet, 0, c).Value = table.Columns[c].ColumnName;
            for (var r = 0; r < table.Rows.Count; r++)
                for (var c = 0; c < table.Columns.Count; c++)
                {
                    var v = table.Rows[r][c];
                    if (v == null || v is DBNull) continue;
                    new Cell(sheet, r + 1, c).Value = v;
                }
        }

        static readonly char[] Forbidden = { '[', ']', ':', '*', '?', '/', '\\' };

        /// <summary>A legal, unused sheet name close to <paramref name="desired"/> (Excel: 31 chars, none of []:*?/\).</summary>
        public static string UniqueSheetName(WorkBook wb, string desired)
        {
            var name = string.Concat((desired ?? "").Where(ch => Array.IndexOf(Forbidden, ch) < 0)).Trim().Trim('\'');
            if (name.Length == 0) name = "Sheet" + (wb.Native.NumberOfSheets + 1);
            if (name.Length > 31) name = name.Substring(0, 31);
            var candidate = name; var i = 2;
            while (wb.Native.GetSheetIndex(candidate) >= 0)
            {
                var suffix = " (" + i++ + ")";
                candidate = name.Substring(0, Math.Min(name.Length, 31 - suffix.Length)) + suffix;
            }
            return candidate;
        }
    }
}
