using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>
    /// NPOI cannot write a table without a header row: it never emits <c>headerRowCount="0"</c> and, on every write,
    /// renames the columns after the cells of the first row. Excel then finds header names that do not match the cells
    /// and asks to repair the file. This remembers the real names before the write and patches the written package.
    /// </summary>
    sealed class HeaderlessTables
    {
        readonly List<(XSSFTable table, string[] names)> _items = new List<(XSSFTable, string[])>();

        public static HeaderlessTables Capture(IWorkbook book)
        {
            var found = new HeaderlessTables();
            for (var i = 0; i < book.NumberOfSheets; i++)
            {
                if (!(book.GetSheetAt(i) is XSSFSheet sheet)) continue;
                foreach (var t in sheet.GetTables())
                {
                    var ct = t.GetCTTable();
                    if (ct.headerRowCount != 0 || ct.tableColumns == null) continue;
                    found._items.Add((t, ct.tableColumns.tableColumn.Select(c => c.name).ToArray()));
                }
            }
            return found;
        }

        public bool Any => _items.Count > 0;

        /// <summary>Puts the names NPOI overwrote back, so a second save starts from the same state.</summary>
        public void Restore()
        {
            foreach (var (table, names) in _items)
            {
                var cols = table.GetCTTable().tableColumns.tableColumn;
                for (var i = 0; i < cols.Count && i < names.Length; i++) cols[i].name = names[i];
            }
        }

        /// <summary>Adds headerRowCount="0" and restores the column names in the table parts of an open package.</summary>
        public void Patch(ZipArchive zip)
        {
            if (!Any) return;
            var byName = _items.ToDictionary(x => x.table.Name, x => x.names);
            foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("xl/tables/", System.StringComparison.Ordinal) && e.FullName.EndsWith(".xml", System.StringComparison.Ordinal)).ToList())
            {
                XDocument doc;
                using (var s = entry.Open()) doc = XDocument.Load(s);
                var root = doc.Root;
                var name = (string)root.Attribute("name");
                if (name == null || !byName.TryGetValue(name, out var names)) continue;

                var ns = root.Name.Namespace;
                root.SetAttributeValue("headerRowCount", "0");
                root.Element(ns + "autoFilter")?.Remove();
                var cols = root.Element(ns + "tableColumns")?.Elements(ns + "tableColumn").ToList() ?? new List<XElement>();
                for (var i = 0; i < cols.Count && i < names.Length; i++) cols[i].SetAttributeValue("name", names[i]);

                var full = entry.FullName;
                entry.Delete();
                using (var s = zip.CreateEntry(full).Open()) doc.Save(s);
            }
        }
    }
}
