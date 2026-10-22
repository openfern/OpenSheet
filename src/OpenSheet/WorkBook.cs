using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal;
using OpenSheet.Internal.Formulas;
using OpenSheet.Metadata;
using OpenSheet.Options;
using XL = NPOI.SS.UserModel;

namespace OpenSheet
{
    /// <summary>A whole spreadsheet file. Usually the first object you create.</summary>
    public partial class WorkBook
    {
        readonly Dictionary<XL.ISheet, WorkSheet> _sheets = new Dictionary<XL.ISheet, WorkSheet>();
        readonly Dictionary<(WorkBook, short), XL.ICellStyle> _imported = new Dictionary<(WorkBook, short), XL.ICellStyle>();
        XL.IFormulaEvaluator _eval;
        XL.DataFormatter _formatter;
        bool _dirty;
        string _filePath;
        string _password;

        // NPOI keeps its function tables in static dictionaries that are not safe to change while another thread parses a formula.
        // Filling them in here, before the first workbook exists, makes sure nobody is reading them yet.
        static WorkBook() { FunctionPack.Initialize(); }

        internal XL.IWorkbook Native { get; private set; }
        internal StyleEditor Styles { get; private set; }
        internal ExcelFileFormat Format { get; private set; }

        internal WorkBook(Loaded loaded)
        {
            Native = loaded.Native;
            Format = loaded.Format;
            Styles = new StyleEditor(Native);
            _filePath = loaded.Path;
            _password = loaded.Password;
            loaded.Populate?.Invoke(this);
        }

        /// <summary>Creates an empty workbook with one sheet named "Sheet1".</summary>
        public WorkBook(ExcelFileFormat fileFormat = ExcelFileFormat.XLSX) : this(Loader.Empty(fileFormat))
        {
            Native.CreateSheet("Sheet1");
        }

        public WorkBook(byte[] spreadSheetBytes) : this(Loader.Open(spreadSheetBytes, null, null, null)) { }
        public WorkBook(Stream spreadSheetStream) : this(Loader.Open(Loader.ReadAll(spreadSheetStream), null, null, null)) { }
        public WorkBook(string spreadSheetFilePath) : this(Loader.OpenFile(spreadSheetFilePath, null)) { }

        /// <summary>Workbook with no sheets (an invalid state for Excel until one is added).</summary>
        internal static WorkBook NewEmpty(ExcelFileFormat fmt) => new WorkBook(Loader.Empty(fmt));

        internal HSSFWorkbook Hssf => (HSSFWorkbook)Native;

        // ---- evaluation, change tracking ----

        bool _edited;

        /// <summary>Marks the workbook as edited so cached formula results are recomputed on next use.</summary>
        internal void Touch() { _dirty = true; _edited = true; }

        // NPOI saves formulas without their results, so readers that do not recalculate (parsers, previewers)
        // would see empty cells. Compute what we can; Excel is told to recalculate on open and stays authoritative.
        void CacheFormulaResults()
        {
            if (!_edited) return;
            var any = false;
            var evaluator = Evaluator;
            using (CellValues.EnterInvariant()) // once for the whole pass, not once per formula
                foreach (var sheet in SheetList())
                    foreach (XL.IRow row in sheet.Native)
                    {
                        if (row == null) continue;
                        foreach (var cell in row)
                        {
                            if (cell.CellType != XL.CellType.Formula) continue;
                            any = true;
                            try { evaluator.EvaluateFormulaCell(cell); }
                            catch (Exception) { /* function NPOI cannot evaluate: leave no cached value, Excel computes it */ }
                        }
                    }
            if (any && Native.NumberOfSheets > 0) Native.GetSheetAt(0).ForceFormulaRecalculation = true;
            _edited = false;
        }

        internal XL.IFormulaEvaluator Evaluator
        {
            get
            {
                if (_eval == null)
                {
                    FunctionPack.Attach(Native);
                    _eval = Native.GetCreationHelper().CreateFormulaEvaluator();
                    _dirty = false;
                }
                else if (_dirty && !AccellerateFormulaEvaluation) { _eval.ClearAllCachedResultValues(); _dirty = false; }
                return _eval;
            }
        }

        internal XL.DataFormatter Formatter => _formatter ?? (_formatter = new XL.DataFormatter(CultureInfo.InvariantCulture));

        // Built from the style table on first use of each style. A style is never edited once cells use it (edits clone it into a new one),
        // so an entry stays right for the life of the workbook.
        NumberStyle[] _numberStyles = new NumberStyle[0];

        /// <summary>The number format of the cell's style, looked up once per style instead of once per cell (NPOI builds a style object per call).</summary>
        internal NumberStyle NumberStyleOf(XL.ICell cell)
        {
            var index = CellValues.StyleIndex(cell);
            var known = Native.NumCellStyles;
            if (index < 0 || index >= known) return NumberStyle.Of(cell);
            if (index >= _numberStyles.Length) Array.Resize(ref _numberStyles, Math.Min(known, Math.Max(index + 1, _numberStyles.Length * 2)));
            return _numberStyles[index] ?? (_numberStyles[index] = NumberStyle.Of(cell));
        }

        /// <summary>
        /// When true, formula results stay cached between edits (faster, less memory) - call <see cref="EvaluateAll"/>
        /// after editing. When false, every edit invalidates the cache. (The spelling follows the established API.)
        /// </summary>
        public bool AccellerateFormulaEvaluation { get; set; }

        /// <summary>Computes every formula and stores the results; formula cells stay formulas.</summary>
        public WorkBook EvaluateAll()
        {
            var evaluator = Evaluator;
            evaluator.ClearAllCachedResultValues();
            // One cell NPOI cannot compute (an unsupported function) must not stop the others, so no EvaluateAll().
            var formulas = new List<XL.ICell>(); // one row at a time, so evaluating cannot disturb the row being walked
            foreach (var sheet in SheetList())
                foreach (XL.IRow row in sheet.Native)
                {
                    if (row == null) continue;
                    formulas.Clear();
                    foreach (var cell in row) if (cell.CellType == XL.CellType.Formula) formulas.Add(cell);
                    foreach (var cell in formulas) CellValues.IsUnsupported(evaluator, cell);
                }
            _dirty = false;
            return this;
        }

        // ---- sheets ----

        internal WorkSheet Wrap(XL.ISheet sheet)
        {
            if (!_sheets.TryGetValue(sheet, out var ws)) _sheets[sheet] = ws = new WorkSheet(this, sheet);
            return ws;
        }

        internal List<WorkSheet> SheetList() => Enumerable.Range(0, Native.NumberOfSheets).Select(i => Wrap(Native.GetSheetAt(i))).ToList();

        public WorksheetsCollection WorkSheets => new WorksheetsCollection(this);

        /// <summary>The first sheet.</summary>
        public WorkSheet DefaultWorkSheet => Native.NumberOfSheets == 0 ? null : Wrap(Native.GetSheetAt(0));

        /// <summary>The sheet with that name (case-insensitive), or null.</summary>
        public WorkSheet GetWorkSheet(string workSheetName)
        {
            var i = Native.GetSheetIndex(workSheetName);
            return i < 0 ? null : Wrap(Native.GetSheetAt(i));
        }

        public WorkSheet CreateWorkSheet(string name)
        {
            var s = Wrap(Native.CreateSheet(name));
            Touch();
            return s;
        }

        public void RemoveWorkSheet(string name)
        {
            var i = Native.GetSheetIndex(name);
            if (i < 0) throw new ArgumentException($"No worksheet named '{name}'.", nameof(name));
            RemoveWorkSheet(i);
        }

        public void RemoveWorkSheet(int index)
        {
            if (index < 0 || index >= Native.NumberOfSheets) throw new ArgumentOutOfRangeException(nameof(index));
            var sheet = Native.GetSheetAt(index);
            _sheets.Remove(sheet);
            SheetParts.Release(sheet);
            Native.RemoveSheetAt(index);
            Touch();
        }

        public void SetSheetPosition(string sheetName, int index)
        {
            if (index < 0 || index >= Native.NumberOfSheets) throw new ArgumentOutOfRangeException(nameof(index));
            Native.SetSheetOrder(sheetName, index);
        }

        public void SetActiveTab(int index) => Native.SetActiveSheet(index);

        public void Close() => Native.Close();

        // ---- identity ----

        /// <summary>Document properties: author, title, dates, custom properties.</summary>
        public IMetadata Metadata => Native is XSSFWorkbook x ? (IMetadata)new XlsxMetadata(x) : new XlsMetadata(Hssf);

        /// <summary>Full path of the file this workbook was loaded from or last saved to; null if it has none.</summary>
        public string FilePath => _filePath;

        /// <summary>When set, saving encrypts the workbook (XLSX only).</summary>
        public string Password { get => _password; set => _password = string.IsNullOrEmpty(value) ? null : value; }

        public bool IsEncrypted
        {
            get => _password != null;
            set
            {
                if (!value) _password = null;
                else if (_password == null) throw new InvalidOperationException("Set Password to encrypt the workbook.");
            }
        }

        // ---- styles shared across workbooks ----

        internal XL.ICellStyle ImportStyle(XL.ICellStyle style, WorkBook from)
        {
            if (ReferenceEquals(from, this)) return style;
            var k = (from, style.Index);
            if (!_imported.TryGetValue(k, out var s))
            {
                if (from.Format == Format) { s = Native.CreateCellStyle(); s.CloneStyleFrom(style); }
                else s = SheetCopier.CrossStyle(from, style, this);
                _imported[k] = s;
            }
            return s;
        }

        // ---- creating ----

        public static WorkBook Create() => new WorkBook(ExcelFileFormat.XLSX);
        public static WorkBook Create(ExcelFileFormat fileFormat) => new WorkBook(fileFormat);
        public static WorkBook Create(CreatingOptions options) => new WorkBook(options?.DefaultFileFormat ?? ExcelFileFormat.XLSX);

        // ---- loading ----

        public static WorkBook Load(string filename) => new WorkBook(Loader.OpenFile(filename, null));
        public static WorkBook Load(string filename, string password) => new WorkBook(Loader.OpenFile(filename, new LoadingOptions { Password = password }));
        public static WorkBook Load(byte[] data) => new WorkBook(Loader.Open(data, null, null, null));
        public static WorkBook Load(byte[] data, string password) => new WorkBook(Loader.Open(data, null, null, new LoadingOptions { Password = password }));
        public static WorkBook Load(Stream stream) => new WorkBook(Loader.Open(Loader.ReadAll(stream), null, null, null));
        public static WorkBook Load(Stream stream, string password) => new WorkBook(Loader.Open(Loader.ReadAll(stream), null, null, new LoadingOptions { Password = password }));

        public static WorkBook Load(string filename, string listDelimiter = null, ExcelFileFormat fileFormat = ExcelFileFormat.XLSX) =>
            new WorkBook(Loader.OpenFile(filename, new LoadingOptions { CsvListDelimiter = listDelimiter, DefaultFileFormat = fileFormat }));

        public static WorkBook Load(string filename, LoadingOptions options) => new WorkBook(Loader.OpenFile(filename, options));

        public static WorkBook Load(DataSet dataset, CreatingOptions options = null)
        {
            if (dataset == null) throw new ArgumentNullException(nameof(dataset));
            var wb = NewEmpty(options?.DefaultFileFormat ?? ExcelFileFormat.XLSX);
            LoadWorkSheetsFromDataSet(dataset, wb);
            if (wb.Native.NumberOfSheets == 0) wb.Native.CreateSheet("Sheet1");
            return wb;
        }

        public static WorkBook LoadExcel(byte[] data) => Load(data);
        public static WorkBook LoadExcel(byte[] data, string password) => Load(data, password);
        public static WorkBook LoadExcel(Stream file) => Load(file);
        public static WorkBook LoadExcel(Stream file, string password) => Load(file, password);
        public static WorkBook LoadExcel(string filename) => Load(filename);
        public static WorkBook LoadExcel(string filename, string password) => Load(filename, password);

        public static WorkBook LoadCSV(string filename, ExcelFileFormat fileFormat = ExcelFileFormat.XLSX, string listDelimiter = null, bool convertFieldValues = true) =>
            new WorkBook(Loader.OpenFile(filename, new LoadingOptions { DefaultFileFormat = fileFormat, CsvListDelimiter = listDelimiter, ConvertFieldValues = convertFieldValues }, forceText: true));

        public static WorkBook Import(string filename, ExcelFileFormat fileFormat = ExcelFileFormat.XLSX) =>
            new WorkBook(Loader.OpenFile(filename, new LoadingOptions { DefaultFileFormat = fileFormat }, forceText: true));

        public static WorkBook FromByteArray(byte[] data) => Load(data);
        public static WorkBook FromByteArray(byte[] data, string password) => Load(data, password);
        public static WorkBook FromStream(Stream stream) => Load(stream);
        public static WorkBook FromStream(Stream stream, string password) => Load(stream, password);

        public static bool TryLoad(string filename, out WorkBook workBook) => Try(() => Load(filename), out workBook);
        public static bool TryLoad(string filename, string password, out WorkBook workBook) => Try(() => Load(filename, password), out workBook);
        public static bool TryLoad(string filename, out WorkBook workBook, string listDelimiter = null, ExcelFileFormat fileFormat = ExcelFileFormat.XLSX) =>
            Try(() => Load(filename, listDelimiter, fileFormat), out workBook);
        public static bool TryLoad(string filename, out WorkBook workBook, LoadingOptions options) => Try(() => Load(filename, options), out workBook);
        public static bool TryLoadExcel(string filename, out WorkBook workBook) => TryLoad(filename, out workBook);
        public static bool TryLoadExcel(string filename, string password, out WorkBook workBook) => TryLoad(filename, password, out workBook);

        static bool Try(Func<WorkBook> load, out WorkBook wb)
        {
            try { wb = load(); return true; }
            catch (Exception) { wb = null; return false; }
        }

        // ---- saving ----

        public WorkBook Save() => Save(null);

        public WorkBook Save(SavingOptions options)
        {
            if (_filePath == null) throw new InvalidOperationException("This workbook has no file path yet; use SaveAs.");
            return SaveAs(_filePath, options);
        }

        public WorkBook SaveAs(string fileName) => SaveAs(fileName, (SavingOptions)null);
        public WorkBook SaveAs(string fileName, string password) => SaveAs(fileName, new SavingOptions { Password = password });
        public WorkBook SaveAsWithCustomDelimiter(string fileName, string listDelimiter = null) => SaveAs(fileName, new SavingOptions { ListDelimiter = listDelimiter });

        public WorkBook SaveAs(string fileName, SavingOptions options) =>
            WriteAs(fileName, Path.GetExtension(fileName ?? ""), options);

        // The format comes from "extension", not the file name, so SaveAsCsv("x.txt") writes CSV to x.txt.
        WorkBook WriteAs(string fileName, string extension, SavingOptions options)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("The file name is empty.", nameof(fileName));
            var password = options?.Password ?? _password;
            var ext = extension.ToLowerInvariant();
            switch (ext)
            {
                case ".xlsx": case ".xlsm": case ".xltx": case ".xltm": File.WriteAllBytes(fileName, Serialize(ExcelFileFormat.XLSX, password)); break;
                case ".xls": case ".xlt": File.WriteAllBytes(fileName, Serialize(ExcelFileFormat.XLS, password)); break;
                case ".csv": WriteSheets(fileName, options?.ListDelimiter ?? ",", TextExport.Csv); break;
                case ".tsv": WriteSheets(fileName, options?.ListDelimiter ?? "\t", TextExport.Csv); break;
                case ".json": File.WriteAllBytes(fileName, TextExport.Json(SheetList())); break;
                case ".xml": File.WriteAllBytes(fileName, TextExport.Xml(SheetList())); break;
                case ".html": case ".htm": File.WriteAllBytes(fileName, TextExport.Bytes(TextExport.Html(this, new HtmlExportOptions()))); break;
                default: throw new NotSupportedException($"Cannot save files with extension '{ext}'. Use .xlsx, .xls, .csv, .tsv, .json, .xml or .html.");
            }
            _filePath = Path.GetFullPath(fileName);
            return this;
        }

        // One file per sheet: the first keeps the requested name, the others get ".<sheet name>" before the extension.
        void WriteSheets(string fileName, string delimiter, Func<WorkSheet, string, byte[]> render)
        {
            var sheets = SheetList();
            var dir = Path.GetDirectoryName(Path.GetFullPath(fileName));
            var stem = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            for (var i = 0; i < sheets.Count; i++)
            {
                var path = i == 0 ? fileName : Path.Combine(dir, stem + "." + Sanitize(sheets[i].Name) + ext);
                File.WriteAllBytes(path, render(sheets[i], delimiter));
            }
        }

        static string Sanitize(string s)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var chars = s.ToCharArray();
            for (var i = 0; i < chars.Length; i++) if (Array.IndexOf(invalid, chars[i]) >= 0) chars[i] = '_';
            return new string(chars);
        }

        public void SaveAsCsv(string filename, string delimiter = null) => WriteAs(filename, ".csv", new SavingOptions { ListDelimiter = delimiter });
        public void SaveAsJson(string filename) => WriteAs(filename, ".json", null);
        public void SaveAsXml(string filename) => WriteAs(filename, ".xml", null);

        internal byte[] Serialize(ExcelFileFormat fmt, string password)
        {
            var source = fmt == Format ? this : ConvertTo(fmt);
            foreach (var sheet in source.SheetList()) sheet.PlotPendingCharts();
            source.CacheFormulaResults();
            byte[] bytes;
            var headerless = fmt == ExcelFileFormat.XLSX ? HeaderlessTables.Capture(source.Native) : null;
            if (fmt == ExcelFileFormat.XLSX) PageSetupFix.Apply(source.Native);
            using (var ms = new MemoryStream())
            {
                try { source.Native.Write(ms, true); }
                finally { headerless?.Restore(); }
                bytes = fmt == ExcelFileFormat.XLSX ? PackageFixes.Apply(ms.ToArray(), headerless) : ms.ToArray();
            }
            if (password == null) return bytes;
            if (fmt != ExcelFileFormat.XLSX) throw new NotSupportedException("Password protection is supported for the XLSX format only.");
            return Crypto.Encrypt(bytes, password);
        }

        // Copies sheets cell by cell, then the defined names and the document properties. Drawings (pictures, charts),
        // data validations and conditional formats are not carried across formats.
        WorkBook ConvertTo(ExcelFileFormat fmt)
        {
            var copy = NewEmpty(fmt);
            foreach (var s in SheetList()) SheetCopier.Copy(s, copy, s.Name);
            foreach (var name in Native.GetAllNames().Where(n => !n.IsFunctionName && !n.IsDeleted && !n.NameName.StartsWith("_xlnm", StringComparison.Ordinal) && !string.IsNullOrEmpty(n.RefersToFormula)))
            {
                var made = copy.Native.CreateName();
                if (name.SheetIndex >= 0) made.SheetIndex = name.SheetIndex;
                made.NameName = name.NameName;
                made.RefersToFormula = name.RefersToFormula;
            }
            MetadataCopier.Copy(Metadata, copy.Metadata);
            return copy;
        }

        internal WorkBook SingleSheetBook(WorkSheet sheet, ExcelFileFormat fmt)
        {
            var wb = NewEmpty(fmt);
            SheetCopier.Copy(sheet, wb, sheet.Name);
            return wb;
        }

        internal void SaveSheet(WorkSheet sheet, string filename)
        {
            var ext = Path.GetExtension(filename).ToLowerInvariant();
            switch (ext)
            {
                case ".csv": File.WriteAllBytes(filename, TextExport.Csv(sheet, ",")); break;
                case ".tsv": File.WriteAllBytes(filename, TextExport.Csv(sheet, "\t")); break;
                case ".json": File.WriteAllBytes(filename, TextExport.Json(sheet)); break;
                case ".xml": File.WriteAllBytes(filename, TextExport.Xml(new[] { sheet })); break;
                case ".xls": SingleSheetBook(sheet, ExcelFileFormat.XLS).SaveAs(filename); break;
                case ".xlsx": SingleSheetBook(sheet, ExcelFileFormat.XLSX).SaveAs(filename); break;
                default: throw new NotSupportedException($"Cannot export a worksheet to '{ext}'.");
            }
        }

        // ---- streams ----

        public MemoryStream ToStream(ExcelFileFormat? format = null) => new MemoryStream(Serialize(format ?? Format, _password));
        public MemoryStream ToXlsStream() => new MemoryStream(Serialize(ExcelFileFormat.XLS, null));
        public MemoryStream ToXlsxStream() => new MemoryStream(Serialize(ExcelFileFormat.XLSX, _password));
        public byte[] ToBinary() => Serialize(Format, _password);
        public byte[] ToByteArray() => ToBinary();

        WorkSheet SheetFor(int? index)
        {
            var s = index.HasValue ? WorkSheets[index.Value] : DefaultWorkSheet;
            return s ?? throw new InvalidOperationException("The workbook has no worksheets.");
        }

        public MemoryStream ToCsvStream(int? sheetIndex = null, string delimiter = null) => TextExport.Stream(TextExport.Csv(SheetFor(sheetIndex), delimiter ?? ","));
        public IEnumerable<MemoryStream> ToCsvStreams(string delimiter = null) => SheetList().Select(s => TextExport.Stream(TextExport.Csv(s, delimiter ?? ","))).ToList();
        public MemoryStream ToTsvStream(int? sheetIndex = null) => TextExport.Stream(TextExport.Csv(SheetFor(sheetIndex), "\t"));
        public IEnumerable<MemoryStream> ToTsvStreams() => SheetList().Select(s => TextExport.Stream(TextExport.Csv(s, "\t"))).ToList();
        public MemoryStream ToJsonStream(int? sheetIndex = null) => TextExport.Stream(TextExport.Json(SheetFor(sheetIndex)));
        public IEnumerable<MemoryStream> ToJsonStreams() => SheetList().Select(s => TextExport.Stream(TextExport.Json(s))).ToList();
        public MemoryStream ToXmlStream(int? sheetIndex = null) => TextExport.Stream(TextExport.Xml(new[] { SheetFor(sheetIndex) }));
        public IEnumerable<MemoryStream> ToXmlStreams() => SheetList().Select(s => TextExport.Stream(TextExport.Xml(new[] { s }))).ToList();

        // ---- HTML ----

        public string ExportToHtmlString(HtmlExportOptions options) => TextExport.Html(this, options ?? new HtmlExportOptions());

        public string ExportToHtmlString(bool outputColumnHeaders = true, bool outputHiddenColumns = false, bool outputHiddenRows = true,
            bool outputLeadingSpacesAsNonBreaking = true, bool outputRowNumbers = true) =>
            ExportToHtmlString(new HtmlExportOptions
            {
                OutputColumnHeaders = outputColumnHeaders, OutputHiddenColumns = outputHiddenColumns, OutputHiddenRows = outputHiddenRows,
                OutputLeadingSpacesAsNonBreaking = outputLeadingSpacesAsNonBreaking, OutputRowNumbers = outputRowNumbers,
            });

        public WorkBook ExportToHtml(string filename, HtmlExportOptions options)
        {
            File.WriteAllBytes(filename, TextExport.Bytes(ExportToHtmlString(options)));
            return this;
        }

        public WorkBook ExportToHtml(string filename) => ExportToHtml(filename, new HtmlExportOptions());

        public WorkBook ExportToHtml(string filename, bool outputColumnHeaders = true, bool outputHiddenColumns = false, bool outputHiddenRows = true,
            bool outputLeadingSpacesAsNonBreaking = true, bool outputRowNumbers = true) =>
            ExportToHtml(filename, new HtmlExportOptions
            {
                OutputColumnHeaders = outputColumnHeaders, OutputHiddenColumns = outputHiddenColumns, OutputHiddenRows = outputHiddenRows,
                OutputLeadingSpacesAsNonBreaking = outputLeadingSpacesAsNonBreaking, OutputRowNumbers = outputRowNumbers,
            });

        // ---- data sets ----

        public DataSet ToDataSet(bool useFirstRowAsColumnNames = false)
        {
            var ds = new DataSet();
            foreach (var s in SheetList()) ds.Tables.Add(TableConverter.ToDataTable(s, useFirstRowAsColumnNames, s.Name));
            return ds;
        }

        public static void LoadWorkSheetsFromDataSet(DataSet dataset, WorkBook workBook)
        {
            if (dataset == null) throw new ArgumentNullException(nameof(dataset));
            if (workBook == null) throw new ArgumentNullException(nameof(workBook));
            foreach (DataTable t in dataset.Tables) workBook.LoadWorkSheet(t);
        }

        /// <summary>Adds a sheet named after the table, with column names in the first row.</summary>
        public WorkSheet LoadWorkSheet(DataTable table)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            var name = TableConverter.UniqueSheetName(this, table.TableName);
            var sheet = CreateWorkSheet(name);
            TableConverter.Fill(sheet, table);
            return sheet;
        }

        // ---- encryption ----

        public MemoryStream Encrypt(string password = null)
        {
            var pw = password ?? _password ?? throw new InvalidOperationException("Pass a password or set the Password property.");
            return new MemoryStream(Crypto.Encrypt(Serialize(ExcelFileFormat.XLSX, null), pw));
        }

        public static MemoryStream Decrypt(Stream stream, string password) => new MemoryStream(Crypto.Decrypt(Loader.ReadAll(stream), password));
    }
}
