namespace OpenSheet
{
    /// <summary>File format used when a workbook is created or saved.</summary>
    public enum ExcelFileFormat
    {
        /// <summary>Office Open XML (.xlsx). The default.</summary>
        XLSX = 1,
        /// <summary>Legacy binary format (.xls).</summary>
        XLS = 2,
    }

    public enum SortOrder { Ascending = 0, Descending = 1 }

    /// <summary>Kind of value held by a <see cref="StreamCell"/>.</summary>
    public enum StreamCellType { Blank = 0, Number = 1, String = 2, Boolean = 3, Date = 4, Error = 5 }
}
