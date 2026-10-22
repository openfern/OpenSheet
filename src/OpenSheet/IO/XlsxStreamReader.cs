using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>
    /// Reads one worksheet of an .xlsx package row by row with XmlReader, without building a node tree per row. Shared strings and number
    /// formats are loaded up front (small); the sheet itself is never held whole. DTDs and external entities are refused.
    /// </summary>
    internal sealed class XlsxStreamReader : IDisposable
    {
        static readonly XmlReaderSettings Safe = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreComments = true, IgnoreProcessingInstructions = true,
        };

        // Parts read whole (workbook, relationships, styles) are small; the cap stops a decompression bomb from filling memory.
        static readonly XmlReaderSettings SafeSmall = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreComments = true, IgnoreProcessingInstructions = true,
            MaxCharactersInDocument = 64L * 1024 * 1024,
        };

        readonly ZipArchive _zip;

        public XlsxStreamReader(Stream zipStream)
        {
            try { _zip = new ZipArchive(zipStream, ZipArchiveMode.Read, false); }
            catch (InvalidDataException) { throw new NotSupportedException("Streaming reads .xlsx only (encrypted and .xls files are not supported)."); }
        }

        public void Dispose() => _zip.Dispose();

        ZipArchiveEntry Entry(string path)
        {
            path = path.TrimStart('/').Replace('\\', '/');
            return _zip.GetEntry(path) ?? _zip.Entries.FirstOrDefault(e => string.Equals(e.FullName, path, StringComparison.OrdinalIgnoreCase));
        }

        XDocument Load(string path)
        {
            var e = Entry(path);
            if (e == null) return null;
            using (var s = e.Open())
            using (var xr = XmlReader.Create(s, SafeSmall)) return XDocument.Load(xr);
        }

        static string RelId(XElement e) => e.Attributes().FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.NamespaceName.EndsWith("relationships", StringComparison.Ordinal))?.Value;

        static string Combine(string baseDir, string target) =>
            target.StartsWith("/") ? target.TrimStart('/') : (baseDir.Length == 0 ? "" : baseDir + "/") + target;

        string _workbookPath;

        string WorkbookPath()
        {
            if (_workbookPath != null) return _workbookPath;
            var rels = Load("_rels/.rels");
            var t = rels?.Descendants().FirstOrDefault(e => e.Name.LocalName == "Relationship" && (e.Attribute("Type")?.Value ?? "").EndsWith("/officeDocument", StringComparison.Ordinal))?.Attribute("Target")?.Value;
            return _workbookPath = (t ?? "xl/workbook.xml").TrimStart('/');
        }

        string WorkbookDir()
        {
            var wbPath = WorkbookPath();
            return wbPath.Contains("/") ? wbPath.Substring(0, wbPath.LastIndexOf('/')) : "";
        }

        /// <summary>Sheets in tab order with the package path of each worksheet part.</summary>
        public List<(string Name, string Path)> Sheets()
        {
            var wbPath = WorkbookPath();
            var dir = WorkbookDir();
            var wb = Load(wbPath) ?? throw new InvalidDataException("The package has no workbook part.");
            var relsPath = (dir.Length == 0 ? "" : dir + "/") + "_rels/" + Path.GetFileName(wbPath) + ".rels";
            var targets = (Load(relsPath)?.Descendants().Where(e => e.Name.LocalName == "Relationship") ?? Enumerable.Empty<XElement>())
                .ToDictionary(e => e.Attribute("Id").Value, e => e.Attribute("Target").Value);
            return wb.Descendants().Where(e => e.Name.LocalName == "sheet")
                .Select(e => (e.Attribute("name").Value, Combine(dir, targets[RelId(e)]))).ToList();
        }

        List<string> SharedStrings()
        {
            var list = new List<string>();
            var entry = Entry(Combine(WorkbookDir(), "sharedStrings.xml"));
            if (entry == null) return list;
            using (var s = entry.Open())
            using (var xr = XmlReader.Create(s, Safe))
            {
                while (!xr.EOF)
                {
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "si") list.Add(ReadText(xr));
                    else xr.Read();
                }
            }
            return list;
        }

        /// <summary>
        /// The text of a rich-text element (a shared string &lt;si&gt; or an inline &lt;is&gt;): its &lt;t&gt; parts joined. Phonetic runs (rPh)
        /// are readings, not part of the text. Leaves the reader after the element.
        /// </summary>
        static string ReadText(XmlReader xr)
        {
            if (xr.IsEmptyElement) { xr.Read(); return ""; }
            var depth = xr.Depth;
            string text = null; List<string> parts = null;
            xr.Read();
            while (xr.Depth > depth)
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "rPh") xr.Skip();
                else if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "t")
                {
                    var part = xr.ReadElementContentAsString();
                    if (text == null) text = part;
                    else (parts ?? (parts = new List<string> { text })).Add(part);
                }
                else xr.Read();
            }
            xr.Read(); // past the end tag
            return parts != null ? string.Concat(parts) : text ?? "";
        }

        bool[] DateStyles()
        {
            var doc = Load(Combine(WorkbookDir(), "styles.xml"));
            if (doc == null) return new bool[0];
            var custom = new Dictionary<int, string>();
            foreach (var e in doc.Descendants().Where(e => e.Name.LocalName == "numFmt"))
                custom[int.Parse(e.Attribute("numFmtId").Value, CultureInfo.InvariantCulture)] = e.Attribute("formatCode").Value;
            var xfs = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "cellXfs");
            if (xfs == null) return new bool[0];
            return xfs.Elements().Where(e => e.Name.LocalName == "xf").Select(xf =>
            {
                var id = int.Parse(xf.Attribute("numFmtId")?.Value ?? "0", CultureInfo.InvariantCulture);
                return XL.DateUtil.IsADateFormat(id, custom.TryGetValue(id, out var code) ? code : null);
            }).ToArray();
        }

        public IEnumerable<StreamRow> Rows(string sheetPath, StreamReadOptions options)
        {
            var entry = Entry(sheetPath) ?? throw new InvalidDataException($"The package has no part '{sheetPath}'.");
            var sst = SharedStrings();
            var dates = DateStyles();
            List<string> header = null; Dictionary<string, int> names = null;
            var rowNum = -1; var width = 4;

            using (var s = entry.Open())
            using (var xr = XmlReader.Create(s, Safe))
            {
                while (!xr.EOF)
                {
                    if (!(xr.NodeType == XmlNodeType.Element && xr.LocalName == "row")) { xr.Read(); continue; }
                    var rAttr = xr.GetAttribute("r");
                    rowNum = rAttr != null ? int.Parse(rAttr, CultureInfo.InvariantCulture) - 1 : rowNum + 1;

                    var cells = ReadCells(xr, sst, dates, width);
                    width = Math.Max(width, cells.Count); // the next row is probably as wide
                    if (options.HasHeaderRow)
                    {
                        if (rowNum < options.HeaderRowIndex) continue;
                        if (rowNum == options.HeaderRowIndex)
                        {
                            header = new List<string>();
                            names = new Dictionary<string, int>();
                            foreach (var c in cells)
                            {
                                while (header.Count <= c.ColumnIndex) header.Add(null);
                                if (string.IsNullOrWhiteSpace(c.Text)) continue;
                                header[c.ColumnIndex] = c.Text;
                                if (!names.ContainsKey(c.Text)) names[c.Text] = c.ColumnIndex;
                            }
                            continue;
                        }
                    }
                    if (cells.Count == 0) continue;
                    yield return new StreamRow(rowNum, cells, header, names);
                }
            }
        }

        /// <summary>The cells of the &lt;row&gt; the reader is on. Leaves the reader after the row.</summary>
        static List<StreamCell> ReadCells(XmlReader xr, List<string> sst, bool[] dates, int expected)
        {
            var cells = new List<StreamCell>(expected);
            if (xr.IsEmptyElement) { xr.Read(); return cells; }
            var depth = xr.Depth;
            var col = -1;
            xr.Read();
            while (xr.Depth > depth)
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "c")
                {
                    var cell = ReadCell(xr, sst, dates, ref col);
                    if (cell != null) cells.Add(cell);
                }
                else xr.Skip();
            }
            xr.Read(); // past </row>
            return cells;
        }

        /// <summary>The cell the reader is on (a &lt;c&gt;), or null when it holds nothing. Leaves the reader after the cell.</summary>
        static StreamCell ReadCell(XmlReader xr, List<string> sst, bool[] dates, ref int col)
        {
            var r = xr.GetAttribute("r");
            col = r != null ? A1.ColumnOf(r) : col + 1;
            var t = xr.GetAttribute("t");
            var style = xr.GetAttribute("s");
            string v = null, inline = null;
            var isFormula = false;

            if (xr.IsEmptyElement) xr.Read();
            else
            {
                var depth = xr.Depth;
                xr.Read();
                while (xr.Depth > depth)
                {
                    if (xr.NodeType != XmlNodeType.Element) { xr.Read(); continue; }
                    switch (xr.LocalName)
                    {
                        case "v": if (v == null) v = xr.ReadElementContentAsString(); else xr.Skip(); break; // the first one counts
                        case "f": isFormula = true; xr.Skip(); break;
                        case "is": if (inline == null) inline = ReadText(xr); else xr.Skip(); break;
                        default: xr.Skip(); break;
                    }
                }
                xr.Read(); // past </c>
            }

            StreamCell cell = null;
            switch (t)
            {
                case "s":
                    if (v != null)
                    {
                        if (!int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) || index < 0 || index >= sst.Count)
                            throw new InvalidDataException($"A cell points at shared string '{v}', which the package does not have.");
                        var str = sst[index];
                        cell = new StreamCell(col, StreamCellType.String, str, str, isFormula);
                    }
                    break;
                case "str":
                    if (v != null) cell = new StreamCell(col, StreamCellType.String, v, v, isFormula);
                    break;
                case "inlineStr":
                    var inl = inline ?? "";
                    cell = new StreamCell(col, StreamCellType.String, inl, inl, isFormula);
                    break;
                case "b":
                    if (v != null) cell = new StreamCell(col, StreamCellType.Boolean, v == "1", v == "1" ? "TRUE" : "FALSE", isFormula);
                    break;
                case "e":
                    if (v != null) cell = new StreamCell(col, StreamCellType.Error, v, v, isFormula);
                    break;
                case "d":
                    if (v != null && DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var iso))
                        cell = new StreamCell(col, StreamCellType.Date, iso, v, isFormula);
                    break;
                default:
                    if (v != null && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                    {
                        var sIdx = int.TryParse(style, NumberStyles.Integer, CultureInfo.InvariantCulture, out var si) ? si : 0;
                        // a date format on a number no date can hold (negative, past year 9999) leaves it a number, as Cell.Value does
                        cell = sIdx >= 0 && sIdx < dates.Length && dates[sIdx] && d >= 0 && d < 2958466 && Serial.TryToDateTime(d, out var when)
                            ? new StreamCell(col, StreamCellType.Date, when, v, isFormula)
                            : new StreamCell(col, StreamCellType.Number, d, v, isFormula);
                    }
                    break;
            }
            if (cell == null && isFormula) cell = new StreamCell(col, StreamCellType.Blank, null, "", true); // formula without a cached result
            return cell;
        }
    }
}
