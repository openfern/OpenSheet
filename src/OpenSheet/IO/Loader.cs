using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml;
using Cysharp.Text;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using OpenSheet.Options;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>What a loader found: the native workbook plus how to finish filling it once the wrapper exists.</summary>
    internal sealed class Loaded
    {
        public XL.IWorkbook Native;
        public ExcelFileFormat Format;
        public string Path;
        public string Password;
        public Action<WorkBook> Populate;
    }

    internal static class Loader
    {
        public static Loaded Empty(ExcelFileFormat fmt) => new Loaded
        {
            Native = fmt == ExcelFileFormat.XLS ? (XL.IWorkbook)new HSSFWorkbook() : new XSSFWorkbook(),
            Format = fmt,
        };

        public static byte[] ReadAll(Stream s)
        {
            if (s == null) throw new ArgumentNullException(nameof(s));
            if (s.CanSeek && s.Length - s.Position <= int.MaxValue)
            {
                // the size is known: one exact buffer instead of a growing MemoryStream plus a final copy
                var data = new byte[(int)(s.Length - s.Position)];
                var read = 0;
                for (int n; read < data.Length && (n = s.Read(data, read, data.Length - read)) > 0;) read += n;
                if (read == data.Length) return data;
                Array.Resize(ref data, read);
                return data;
            }
            using (var ms = new MemoryStream())
            {
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }

        public static Loaded OpenFile(string path, LoadingOptions opts, bool forceText = false)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("The file name is empty.", nameof(path));
            if (!File.Exists(path)) throw new FileNotFoundException("The workbook file was not found.", path);
            return Open(File.ReadAllBytes(path), path, null, opts, forceText);
        }

        public static Loaded Open(byte[] data, string path, string extensionHint, LoadingOptions opts, bool forceText = false)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            opts = opts ?? new LoadingOptions();
            var ext = (extensionHint ?? Path.GetExtension(path ?? "")).ToLowerInvariant();
            Loaded r;
            if (!forceText && IsZip(data)) r = new Loaded { Native = new XSSFWorkbook(new MemoryStream(data)), Format = ExcelFileFormat.XLSX };
            else if (!forceText && IsOle(data)) r = Ole(data, opts.Password);
            else r = Text(data, ext, opts);
            r.Path = path == null ? null : System.IO.Path.GetFullPath(path);
            return r;
        }

        static bool IsZip(byte[] d) => d.Length > 3 && d[0] == 0x50 && d[1] == 0x4B && (d[2] == 3 || d[2] == 5);
        static bool IsOle(byte[] d) => d.Length > 8 && d[0] == 0xD0 && d[1] == 0xCF && d[2] == 0x11 && d[3] == 0xE0 && d[4] == 0xA1 && d[5] == 0xB1;

        static Loaded Ole(byte[] data, string password)
        {
            var fs = new POIFSFileSystem(new MemoryStream(data));
            try
            {
                if (fs.Root.HasEntry("EncryptedPackage"))
                {
                    var xlsx = Crypto.Decrypt(data, password);
                    return new Loaded { Native = new XSSFWorkbook(new MemoryStream(xlsx)), Format = ExcelFileFormat.XLSX, Password = password };
                }
            }
            finally { fs.Close(); }
            return new Loaded { Native = new HSSFWorkbook(new MemoryStream(data)), Format = ExcelFileFormat.XLS };
        }

        // ---- text formats ----

        static string Decode(byte[] d)
        {
            if (d.Length >= 3 && d[0] == 0xEF && d[1] == 0xBB && d[2] == 0xBF) return Encoding.UTF8.GetString(d, 3, d.Length - 3);
            if (d.Length >= 2 && d[0] == 0xFF && d[1] == 0xFE) return Encoding.Unicode.GetString(d, 2, d.Length - 2);
            if (d.Length >= 2 && d[0] == 0xFE && d[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(d, 2, d.Length - 2);
            try { return new UTF8Encoding(false, true).GetString(d); }
            catch (DecoderFallbackException) { return Encoding.GetEncoding("iso-8859-1").GetString(d); } // ANSI files exported by older tools
        }

        static Loaded Text(byte[] data, string ext, LoadingOptions opts)
        {
            var text = Decode(data);
            var head = text.TrimStart();
            Action<WorkBook> fill;
            if (ext == ".json" || (ext != ".csv" && ext != ".tsv" && (head.StartsWith("{") || head.StartsWith("["))))
                fill = wb => TextImport.Json(wb, text);
            else if (ext == ".xml" || (ext != ".csv" && ext != ".tsv" && head.StartsWith("<")))
                fill = wb => TextImport.Xml(wb, text);
            else
            {
                var delimiter = opts.CsvListDelimiter ?? (ext == ".tsv" ? "\t" : TextImport.GuessDelimiter(text));
                fill = wb => TextImport.Csv(wb, text, delimiter, opts.CsvIgnoreQuotes, opts.ConvertFieldValues);
            }
            return new Loaded { Native = Empty(opts.DefaultFileFormat).Native, Format = opts.DefaultFileFormat, Populate = fill };
        }
    }

    internal static class Crypto
    {
        public static byte[] Encrypt(byte[] xlsx, string password)
        {
            if (string.IsNullOrEmpty(password)) throw new ArgumentException("The password is empty.", nameof(password));
            var fs = new POIFSFileSystem();
            try
            {
                // NPOI's default is AES-128 with SHA-1; ask for AES-256 with SHA-512 instead.
                var info = new NPOI.POIFS.Crypt.EncryptionInfo(NPOI.POIFS.Crypt.EncryptionMode.Agile,
                    NPOI.POIFS.Crypt.CipherAlgorithm.aes256, NPOI.POIFS.Crypt.HashAlgorithm.sha512, 256, 16, NPOI.POIFS.Crypt.ChainingMode.cbc);
                var enc = info.Encryptor;
                enc.ConfirmPassword(password);
                using (var os = enc.GetDataStream(fs)) os.Write(xlsx, 0, xlsx.Length);
                using (var ms = new MemoryStream())
                {
                    fs.WriteFileSystem(ms);
                    return ms.ToArray();
                }
            }
            finally { fs.Close(); }
        }

        public static byte[] Decrypt(byte[] data, string password)
        {
            var fs = new POIFSFileSystem(new MemoryStream(data));
            try
            {
                var info = new NPOI.POIFS.Crypt.EncryptionInfo(fs);
                var dec = NPOI.POIFS.Crypt.Decryptor.GetInstance(info);
                if (!dec.VerifyPassword(password ?? NPOI.POIFS.Crypt.Decryptor.DEFAULT_PASSWORD))
                    throw new InvalidDataException("The workbook is encrypted and the password is missing or wrong.");
                using (var s = dec.GetDataStream(fs))
                using (var ms = new MemoryStream())
                {
                    // NPOI's stream signals the end with -1 (Java style); Stream.CopyTo only stops on 0.
                    var buf = new byte[8192];
                    int n;
                    while ((n = s.Read(buf, 0, buf.Length)) > 0) ms.Write(buf, 0, n);
                    return ms.ToArray();
                }
            }
            finally { fs.Close(); }
        }
    }

    internal static class TextImport
    {
        static readonly string[] IsoFormats = { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ssK", "yyyy-MM-ddTHH:mm:ss.FFFFFFFK" };

        /// <summary>Picks ',', ';', tab or '|' by what the first non-empty line uses most outside quotes.</summary>
        public static string GuessDelimiter(string text)
        {
            var best = ','; var bestCount = 0;
            var line = FirstLine(text);
            foreach (var d in new[] { ',', ';', '\t', '|' })
            {
                var n = 0; var q = false;
                foreach (var ch in line) { if (ch == '"') q = !q; else if (ch == d && !q) n++; }
                if (n > bestCount) { best = d; bestCount = n; }
            }
            return best.ToString();
        }

        /// <summary>The first line that holds more than white space, found without splitting the whole text.</summary>
        static string FirstLine(string text)
        {
            for (var start = 0; start < text.Length;)
            {
                var end = text.IndexOf('\n', start);
                if (end < 0) end = text.Length;
                for (var i = start; i < end; i++)
                    if (!char.IsWhiteSpace(text[i])) return text.Substring(start, end - start);
                start = end + 1;
            }
            return "";
        }

        /// <summary>Calls <paramref name="onRow"/> for each non-empty row. The list is reused: copy what you keep.</summary>
        public static void ParseCsv(string text, string delimiter, bool ignoreQuotes, Action<List<string>> onRow)
        {
            var row = new List<string>();
            var sb = ZString.CreateStringBuilder();
            try
            {
                var single = delimiter.Length == 1;
                var d0 = delimiter[0];
                var inQuotes = false; var fieldStarted = false;
                for (var i = 0; i < text.Length; i++)
                {
                    var ch = text[i];
                    if (inQuotes)
                    {
                        if (ch == '"' && i + 1 < text.Length && text[i + 1] == '"') { sb.Append('"'); i++; }
                        else if (ch == '"') inQuotes = false;
                        else sb.Append(ch);
                    }
                    else if (ch == '"' && !ignoreQuotes && sb.Length == 0) { inQuotes = true; fieldStarted = true; }
                    else if (single ? ch == d0 : string.CompareOrdinal(text, i, delimiter, 0, delimiter.Length) == 0)
                    {
                        row.Add(sb.ToString()); sb.Clear(); fieldStarted = false; i += delimiter.Length - 1;
                    }
                    else if (ch == '\r' || ch == '\n')
                    {
                        if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                        if (row.Count > 0 || sb.Length > 0 || fieldStarted) { row.Add(sb.ToString()); onRow(row); }
                        row.Clear(); sb.Clear(); fieldStarted = false;
                    }
                    else sb.Append(ch);
                }
                if (row.Count > 0 || sb.Length > 0 || fieldStarted) { row.Add(sb.ToString()); onRow(row); }
            }
            finally { sb.Dispose(); }
        }

        /// <summary>Number, boolean or ISO date when the text looks like one; otherwise the text. Leading zeros stay text (ZIP codes, IDs).</summary>
        public static object Typed(string s)
        {
            var t = s.Trim();
            if (t.Length == 0) return s;
            if (bool.TryParse(t, out var b)) return b;
            var leadingZero = t.Length > 1 && t[0] == '0' && t[1] != '.' && t[1] != ',';
            if (!leadingZero && double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) && !double.IsNaN(d) && !double.IsInfinity(d)) return d;
            if (t.Length >= 10 && char.IsDigit(t[0]) && DateTime.TryParseExact(t, IsoFormats, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dt)) return dt;
            return s;
        }

        public static void Csv(WorkBook wb, string text, string delimiter, bool ignoreQuotes, bool convert)
        {
            var sheet = wb.Native.NumberOfSheets == 0 ? wb.CreateWorkSheet("Sheet1") : wb.DefaultWorkSheet;
            var r = 0;
            ParseCsv(text, delimiter, ignoreQuotes, fields =>
            {
                XL.IRow row = null;
                for (var c = 0; c < fields.Count; c++)
                {
                    var f = fields[c];
                    if (f.Length == 0) continue;
                    var v = convert ? Typed(f) : f;
                    if (CellValues.IsClearing(v)) continue; // blank text removes a cell; the sheet is new, so there is none to remove
                    if (row == null) row = sheet.Native.GetRow(r) ?? sheet.Native.CreateRow(r);
                    CellValues.Write(wb, row.GetCell(c) ?? row.CreateCell(c), v);
                }
                r++;
            });
            wb.Touch();
        }

        public static void Json(WorkBook wb, string text)
        {
            using (var doc = JsonDocument.Parse(text))
            {
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.Object)
                    foreach (var p in root.EnumerateObject()) JsonSheet(wb, p.Name, p.Value);
                else JsonSheet(wb, "Sheet1", root);
            }
            if (wb.Native.NumberOfSheets == 0) wb.CreateWorkSheet("Sheet1");
        }

        static void JsonSheet(WorkBook wb, string name, JsonElement rows)
        {
            if (rows.ValueKind != JsonValueKind.Array) throw new InvalidDataException("JSON sheets must be arrays of rows.");
            var sheet = wb.CreateWorkSheet(TableConverter.UniqueSheetName(wb, name));
            var list = rows.EnumerateArray().ToList();
            var keys = new List<string>();
            if (list.Count > 0 && list[0].ValueKind == JsonValueKind.Object)
            {
                foreach (var o in list.Where(e => e.ValueKind == JsonValueKind.Object))
                    foreach (var p in o.EnumerateObject()) if (!keys.Contains(p.Name)) keys.Add(p.Name);
                for (var c = 0; c < keys.Count; c++) new Cell(sheet, 0, c).Value = keys[c];
            }
            var r = keys.Count > 0 ? 1 : 0;
            foreach (var item in list)
            {
                if (item.ValueKind == JsonValueKind.Object)
                    foreach (var p in item.EnumerateObject()) JsonCell(new Cell(sheet, r, keys.IndexOf(p.Name)), p.Value);
                else if (item.ValueKind == JsonValueKind.Array)
                {
                    var c = 0;
                    foreach (var v in item.EnumerateArray()) JsonCell(new Cell(sheet, r, c++), v);
                }
                else JsonCell(new Cell(sheet, r, 0), item);
                r++;
            }
        }

        static void JsonCell(Cell cell, JsonElement v)
        {
            switch (v.ValueKind)
            {
                case JsonValueKind.Number: cell.Value = v.GetDouble(); break;
                case JsonValueKind.True: cell.Value = true; break;
                case JsonValueKind.False: cell.Value = false; break;
                case JsonValueKind.String: cell.Value = v.GetString(); break;
                case JsonValueKind.Null: case JsonValueKind.Undefined: break;
                default: cell.Value = v.GetRawText(); break; // nested objects/arrays are kept as JSON text
            }
        }

        public static void Xml(WorkBook wb, string text)
        {
            // The files OpenSheet writes have no DTD; refusing one closes external entities and entity-expansion bombs ("billion laughs").
            var doc = new XmlDocument { XmlResolver = null };
            using (var reader = XmlReader.Create(new StringReader(text), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                doc.Load(reader);
            foreach (XmlElement s in doc.SelectNodes("/workbook/sheet"))
            {
                var sheet = wb.CreateWorkSheet(TableConverter.UniqueSheetName(wb, s.GetAttribute("name")));
                foreach (XmlElement c in s.SelectNodes("row/cell"))
                {
                    A1.ParseCell(c.GetAttribute("ref"), out var r, out var col);
                    var cell = new Cell(sheet, r, col);
                    var v = c.InnerText;
                    switch (c.GetAttribute("type"))
                    {
                        case "number": cell.Value = double.Parse(v, CultureInfo.InvariantCulture); break;
                        case "boolean": cell.Value = bool.Parse(v); break;
                        case "date": cell.Value = DateTime.Parse(v, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind); break;
                        default: cell.Value = v; break;
                    }
                }
            }
            if (wb.Native.NumberOfSheets == 0) wb.CreateWorkSheet("Sheet1");
        }
    }
}
