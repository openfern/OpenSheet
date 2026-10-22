using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenSheet.Internal;

namespace OpenSheet
{
    public partial class WorkBook
    {
        // Forward-only reading of .xlsx worksheets. Nothing is opened until the sequence is enumerated.

        public static IEnumerable<StreamRow> StreamRows(string path, int sheetIndex = 0) => StreamRows(path, new StreamReadOptions { SheetIndex = sheetIndex });
        public static IEnumerable<StreamRow> StreamRows(string path, string sheetName) => StreamRows(path, new StreamReadOptions { SheetName = sheetName });
        public static IEnumerable<StreamRow> StreamRows(Stream stream, int sheetIndex = 0) => StreamRows(stream, new StreamReadOptions { SheetIndex = sheetIndex });
        public static IEnumerable<StreamRow> StreamRows(Stream stream, string sheetName) => StreamRows(stream, new StreamReadOptions { SheetName = sheetName });
        public static IEnumerable<StreamRow> StreamRows(byte[] data, int sheetIndex = 0) => StreamRows(data, new StreamReadOptions { SheetIndex = sheetIndex });
        public static IEnumerable<StreamRow> StreamRows(byte[] data, string sheetName) => StreamRows(data, new StreamReadOptions { SheetName = sheetName });

        public static IEnumerable<StreamRow> StreamRows(string path, StreamReadOptions options)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("The file name is empty.", nameof(path));
            if (options == null) throw new ArgumentNullException(nameof(options));
            return OpenRows(() => new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.SequentialScan), options);
        }

        /// <summary>The stream is copied to memory first so the caller keeps ownership of it; use the path overload for low memory.</summary>
        public static IEnumerable<StreamRow> StreamRows(Stream stream, StreamReadOptions options)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (options == null) throw new ArgumentNullException(nameof(options));
            var data = Loader.ReadAll(stream);
            return OpenRows(() => new MemoryStream(data, false), options);
        }

        public static IEnumerable<StreamRow> StreamRows(byte[] data, StreamReadOptions options)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (options == null) throw new ArgumentNullException(nameof(options));
            return OpenRows(() => new MemoryStream(data, false), options);
        }

        public static IReadOnlyList<string> GetSheetNames(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("The file name is empty.", nameof(path));
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) return Names(fs);
        }

        public static IReadOnlyList<string> GetSheetNames(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            return Names(new MemoryStream(Loader.ReadAll(stream), false));
        }

        public static IReadOnlyList<string> GetSheetNames(byte[] data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            return Names(new MemoryStream(data, false));
        }

        static IReadOnlyList<string> Names(Stream s)
        {
            using (var pkg = new XlsxStreamReader(s)) return pkg.Sheets().Select(x => x.Name).ToList();
        }

        static IEnumerable<StreamRow> OpenRows(Func<Stream> open, StreamReadOptions options)
        {
            using (var pkg = new XlsxStreamReader(open()))
            {
                var sheets = pkg.Sheets();
                int i;
                if (options.SheetName != null)
                {
                    i = sheets.FindIndex(s => string.Equals(s.Name, options.SheetName, StringComparison.OrdinalIgnoreCase));
                    if (i < 0) throw new ArgumentException($"No worksheet named '{options.SheetName}'.", nameof(options));
                }
                else
                {
                    i = options.SheetIndex;
                    if (i < 0 || i >= sheets.Count) throw new ArgumentOutOfRangeException(nameof(options), $"The workbook has {sheets.Count} worksheet(s).");
                }
                foreach (var row in pkg.Rows(sheets[i].Path, options)) yield return row;
            }
        }
    }
}
