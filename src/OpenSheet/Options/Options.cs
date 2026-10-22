namespace OpenSheet.Options
{
    public enum HiddenState { Visible = 0, Hidden = 1, VeryHidden = 2 }

    public enum PanePosition { LowerRight = 0, UpperRight = 1, LowerLeft = 2, UpperLeft = 3 }

    public class CreatingOptions
    {
        public CreatingOptions() { }
        public ExcelFileFormat DefaultFileFormat { get; set; } = ExcelFileFormat.XLSX;
    }

    public class LoadingOptions
    {
        public LoadingOptions() { }
        /// <summary>Separator for CSV input. Null means: comma, or tab for ".tsv".</summary>
        public string CsvListDelimiter { get; set; }
        /// <summary>When true, quote characters in CSV input are ordinary text.</summary>
        public bool CsvIgnoreQuotes { get; set; }
        /// <summary>When true, CSV/JSON/XML text that looks like a number, boolean or date is stored as that type.</summary>
        public bool ConvertFieldValues { get; set; } = true;
        public ExcelFileFormat DefaultFileFormat { get; set; } = ExcelFileFormat.XLSX;
        public string Password { get; set; }
    }

    public class SavingOptions
    {
        public SavingOptions() { }
        /// <summary>Separator for CSV output. Null means: comma, or tab for ".tsv".</summary>
        public string ListDelimiter { get; set; }
        /// <summary>Encrypts the saved workbook (XLSX only).</summary>
        public string Password { get; set; }
    }

    public class HtmlExportOptions
    {
        public HtmlExportOptions() { }
        public bool OutputColumnHeaders { get; set; } = true;
        public bool OutputHiddenColumns { get; set; }
        public bool OutputHiddenRows { get; set; } = true;
        public bool OutputLeadingSpacesAsNonBreaking { get; set; } = true;
        public bool OutputRowNumbers { get; set; } = true;
    }
}
