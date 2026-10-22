using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Xml;
using Cysharp.Text;
using OpenSheet.Options;

namespace OpenSheet.Internal
{
    /// <summary>Plain-text renderings of sheets: CSV/TSV, JSON, XML, HTML. All output is UTF-8 without BOM.</summary>
    internal static class TextExport
    {
        static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
        static readonly char[] CsvSpecials = { '"', '\r', '\n' };

        public static byte[] Bytes(string s) => Utf8.GetBytes(s);
        public static MemoryStream Stream(byte[] bytes) => new MemoryStream(bytes);

        // ---- CSV / TSV ----

        /// <summary>
        /// Text as displayed in Excel (formats applied), RFC 4180 quoting, CRLF line ends.
        /// The block starts at the first used row/column, like a spreadsheet's used range.
        /// </summary>
        public static byte[] Csv(WorkSheet s, string delimiter)
        {
            if (string.IsNullOrEmpty(delimiter)) throw new ArgumentException("The delimiter is empty.", nameof(delimiter));
            if (s.PhysicalRowCount == 0) return Array.Empty<byte>();
            var a = s.RangeAddress;
            var sb = ZString.CreateUtf8StringBuilder(); // encodes as it goes: no UTF-16 string, no second copy to bytes
            try
            {
                for (var r = a.FirstRow; r <= a.LastRow; r++)
                {
                    var row = s.Native.GetRow(r);
                    for (var c = a.FirstColumn; c <= a.LastColumn; c++)
                    {
                        if (c > a.FirstColumn) sb.Append(delimiter);
                        Field(ref sb, Display.Formatted(s.Book, row?.GetCell(c)), delimiter);
                    }
                    sb.Append("\r\n");
                }
                return sb.AsSpan().ToArray();
            }
            finally { sb.Dispose(); }
        }

        static void Field(ref Utf8ValueStringBuilder sb, string v, string delimiter)
        {
            var quote = v.IndexOfAny(CsvSpecials) >= 0 || (delimiter.Length == 1 ? v.IndexOf(delimiter[0]) >= 0 : v.Contains(delimiter));
            if (!quote) { sb.Append(v); return; }
            sb.Append('"');
            var from = 0;
            for (var q = v.IndexOf('"'); q >= 0; q = v.IndexOf('"', from))
            {
                sb.Append(v, from, q + 1 - from); // up to and including the quote, which is then doubled; never splits a surrogate pair
                sb.Append('"');
                from = q + 1;
            }
            sb.Append(v, from, v.Length - from);
            sb.Append('"');
        }

        // ---- JSON ----

        /// <summary>One sheet as an array of rows; each row is an array of values (null for empty cells).</summary>
        public static byte[] Json(WorkSheet s) => Write(w => Rows(w, s));

        /// <summary>Several sheets as an object: { "SheetName": [ rows ] }.</summary>
        public static byte[] Json(IEnumerable<WorkSheet> sheets) => Write(w =>
        {
            w.WriteStartObject();
            foreach (var s in sheets) { w.WritePropertyName(s.Name); Rows(w, s); }
            w.WriteEndObject();
        });

        static byte[] Write(Action<Utf8JsonWriter> body)
        {
            using (var ms = new MemoryStream())
            {
                using (var w = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true })) body(w);
                return ms.ToArray();
            }
        }

        static void Rows(Utf8JsonWriter w, WorkSheet s)
        {
            w.WriteStartArray();
            if (s.PhysicalRowCount > 0)
            {
                var a = s.RangeAddress;
                for (var r = a.FirstRow; r <= a.LastRow; r++)
                {
                    var row = s.Native.GetRow(r);
                    w.WriteStartArray();
                    for (var c = a.FirstColumn; c <= a.LastColumn; c++)
                    {
                        var v = CellValues.Read(s.Book, row?.GetCell(c), out var k);
                        switch (k)
                        {
                            case Kind.Number: w.WriteNumberValue((double)v); break;
                            case Kind.Bool: w.WriteBooleanValue((bool)v); break;
                            case Kind.Date: w.WriteStringValue(((DateTime)v).ToString("s", CultureInfo.InvariantCulture)); break;
                            case Kind.Text: case Kind.Error: w.WriteStringValue((string)v); break;
                            default: w.WriteNullValue(); break;
                        }
                    }
                    w.WriteEndArray();
                }
            }
            w.WriteEndArray();
        }

        // ---- XML ----

        /// <summary>&lt;workbook&gt;&lt;sheet name&gt;&lt;row index&gt;&lt;cell ref type&gt;value. Only cells that hold a value are written.</summary>
        public static byte[] Xml(IEnumerable<WorkSheet> sheets)
        {
            // written to a stream, so the declaration says utf-8 like the bytes do (a StringBuilder target declares utf-16)
            using (var ms = new MemoryStream())
            {
                using (var w = XmlWriter.Create(ms, new XmlWriterSettings { Indent = true, OmitXmlDeclaration = false, Encoding = Utf8 }))
                {
                    w.WriteStartElement("workbook");
                    foreach (var s in sheets)
                    {
                        w.WriteStartElement("sheet");
                        w.WriteAttributeString("name", s.Name);
                        if (s.PhysicalRowCount > 0)
                        {
                            var a = s.RangeAddress;
                            for (var r = a.FirstRow; r <= a.LastRow; r++)
                            {
                                var row = s.Native.GetRow(r);
                                var open = false;
                                for (var c = a.FirstColumn; c <= a.LastColumn; c++)
                                {
                                    var v = CellValues.Read(s.Book, row?.GetCell(c), out var k);
                                    if (k == Kind.Empty) continue;
                                    if (!open) { w.WriteStartElement("row"); w.WriteAttributeString("index", (r + 1).ToString(CultureInfo.InvariantCulture)); open = true; }
                                    w.WriteStartElement("cell");
                                    w.WriteAttributeString("ref", A1.Cell(r, c));
                                    w.WriteAttributeString("type", k == Kind.Number ? "number" : k == Kind.Bool ? "boolean" : k == Kind.Date ? "date" : k == Kind.Error ? "error" : "string");
                                    w.WriteString(k == Kind.Date ? ((DateTime)v).ToString("s", CultureInfo.InvariantCulture)
                                        : k == Kind.Number ? ((double)v).ToString("R", CultureInfo.InvariantCulture)
                                        : k == Kind.Bool ? ((bool)v ? "true" : "false") : (string)v);
                                    w.WriteEndElement();
                                }
                                if (open) w.WriteEndElement();
                            }
                        }
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                }
                return ms.ToArray();
            }
        }

        // ---- HTML ----

        public static string Html(WorkBook wb, HtmlExportOptions o)
        {
            var sb = ZString.CreateStringBuilder();
            try
            {
                sb.Append("<!DOCTYPE html>\n<html>\n<head><meta charset=\"utf-8\"><title>Workbook</title></head>\n<body>\n");
                foreach (var s in wb.SheetList())
                {
                    sb.Append("<h2>"); sb.Append(WebUtility.HtmlEncode(s.Name)); sb.Append("</h2>\n<table border=\"1\" cellspacing=\"0\">\n");
                    var a = s.RangeAddress;
                    var cols = new List<int>(a.ColumnsCount);
                    for (var c = a.FirstColumn; c <= a.LastColumn; c++)
                        if (o.OutputHiddenColumns || !s.Native.IsColumnHidden(c)) cols.Add(c);
                    if (o.OutputColumnHeaders)
                    {
                        sb.Append("<tr>");
                        if (o.OutputRowNumbers) sb.Append("<th></th>");
                        foreach (var c in cols) { sb.Append("<th>"); sb.Append(A1.Column(c)); sb.Append("</th>"); }
                        sb.Append("</tr>\n");
                    }
                    var any = s.PhysicalRowCount > 0;
                    for (var r = a.FirstRow; r <= a.LastRow && any; r++)
                    {
                        var row = s.Native.GetRow(r);
                        if (!o.OutputHiddenRows && (row?.ZeroHeight ?? false)) continue;
                        sb.Append("<tr>");
                        if (o.OutputRowNumbers) { sb.Append("<th>"); sb.Append(r + 1); sb.Append("</th>"); }
                        foreach (var c in cols)
                        {
                            var text = WebUtility.HtmlEncode(Display.Formatted(wb, row?.GetCell(c)));
                            sb.Append("<td>");
                            var lead = 0;
                            if (o.OutputLeadingSpacesAsNonBreaking) while (lead < text.Length && text[lead] == ' ') { sb.Append("&nbsp;"); lead++; }
                            sb.Append(text, lead, text.Length - lead);
                            sb.Append("</td>");
                        }
                        sb.Append("</tr>\n");
                    }
                    sb.Append("</table>\n");
                }
                sb.Append("</body>\n</html>\n");
                return sb.ToString();
            }
            finally { sb.Dispose(); }
        }
    }
}
