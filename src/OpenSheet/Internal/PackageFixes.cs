using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OpenSheet.Internal
{
    /// <summary>
    /// Corrections applied to the finished XLSX package, for things NPOI writes in a form Excel does not accept as intended.
    /// One pass over the zip: tables without a header row, the shape type of comment boxes, and relationships that point at parts which
    /// are no longer in the file (what NPOI leaves behind when a sheet that had charts or pictures is deleted and another sheet takes its place).
    /// </summary>
    internal static class PackageFixes
    {
        // Loading a file that already has comments and saving it again turns the VML shape type id "_x0000_t202" into "202",
        // so every comment box points at a type that does not exist and Excel adds a stray shape.
        static readonly Regex ShapeTypeId = new Regex(@"(<v:shapetype\b[^>]*?\bid="")202("")", RegexOptions.CultureInvariant);

        public static byte[] Apply(byte[] xlsx, HeaderlessTables headerless)
        {
            var needsTables = headerless != null && headerless.Any;
            bool needsVml, needsRelationships;
            using (var probe = new ZipArchive(new MemoryStream(xlsx), ZipArchiveMode.Read))
            {
                needsVml = probe.Entries.Any(e => IsVml(e.FullName));
                needsRelationships = Stale(probe).Count > 0;
            }
            if (!needsTables && !needsVml && !needsRelationships) return xlsx;

            using (var ms = new MemoryStream())
            {
                ms.Write(xlsx, 0, xlsx.Length);
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, true))
                {
                    if (needsTables) headerless.Patch(zip);
                    if (needsVml) FixVml(zip);
                    if (needsRelationships) DropStale(zip);
                }
                return ms.ToArray();
            }
        }

        static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>
        /// Relationships that must go, per relationship file: those whose target is not in the zip, and exact repeats (same type, same target)
        /// of which the owning part refers to none, or to another copy. A worksheet may have only one drawing relationship, for instance.
        /// </summary>
        static System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>> Stale(ZipArchive zip)
        {
            var names = new System.Collections.Generic.HashSet<string>(zip.Entries.Select(e => e.FullName), StringComparer.OrdinalIgnoreCase);
            var stale = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>();
            foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
            {
                XDocument doc;
                using (var s = entry.Open()) doc = XDocument.Load(s);
                if (doc.Root == null) continue;
                var relationships = doc.Root.Elements(Rel + "Relationship").ToList();
                var drop = relationships.Where(r => IsDangling(entry.FullName, r, names)).ToList();

                var owner = OwnerOf(entry.FullName);
                string ownerText = null; var ownerRead = false;
                foreach (var group in relationships.Except(drop).GroupBy(r => (string)r.Attribute("Type") + "|" + Target(entry.FullName, r)).Where(g => g.Count() > 1))
                {
                    // the owning part can be a whole worksheet: read it only when there is a repeat to settle
                    if (!ownerRead)
                    {
                        ownerRead = true;
                        if (owner != null && names.Contains(owner))
                            using (var reader = new StreamReader(zip.GetEntry(owner).Open())) ownerText = reader.ReadToEnd();
                    }
                    var keep = group.FirstOrDefault(r => ownerText != null && ownerText.Contains("\"" + (string)r.Attribute("Id") + "\"")) ?? group.First();
                    drop.AddRange(group.Where(r => r != keep));
                }
                if (drop.Count > 0) stale[entry.FullName] = drop.Select(r => (string)r.Attribute("Id")).ToList();
            }
            return stale;
        }

        /// <summary>"xl/worksheets/_rels/sheet1.xml.rels" belongs to "xl/worksheets/sheet1.xml"; the package's own "_rels/.rels" to nothing.</summary>
        static string OwnerOf(string relsEntry)
        {
            var i = relsEntry.LastIndexOf("/_rels/", StringComparison.Ordinal);
            if (i < 0) return null;
            return relsEntry.Substring(0, i + 1) + relsEntry.Substring(i + "/_rels/".Length, relsEntry.Length - i - "/_rels/".Length - ".rels".Length);
        }

        /// <summary>The target as a zip entry name; null for an external one.</summary>
        static string Target(string relsEntry, XElement relationship)
        {
            if (string.Equals((string)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) return null;
            var target = (string)relationship.Attribute("Target");
            if (string.IsNullOrEmpty(target)) return null;
            if (target.StartsWith("/", StringComparison.Ordinal)) return Uri.UnescapeDataString(target.Substring(1));
            var dir = relsEntry.Contains("/_rels/") ? relsEntry.Substring(0, relsEntry.LastIndexOf("/_rels/", StringComparison.Ordinal)) : "";
            var parts = new System.Collections.Generic.List<string>(dir.Length == 0 ? new string[0] : dir.Split('/'));
            foreach (var seg in target.Split('/'))
            {
                if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else if (seg != "." && seg.Length > 0) parts.Add(seg);
            }
            return Uri.UnescapeDataString(string.Join("/", parts));
        }

        static bool IsDangling(string relsEntry, XElement relationship, System.Collections.Generic.HashSet<string> names)
        {
            var path = Target(relsEntry, relationship);
            return path != null && !names.Contains(path);
        }

        static void DropStale(ZipArchive zip)
        {
            foreach (var pair in Stale(zip))
            {
                var entry = zip.GetEntry(pair.Key);
                XDocument doc;
                using (var s = entry.Open()) doc = XDocument.Load(s);
                doc.Root.Elements(Rel + "Relationship").Where(r => pair.Value.Contains((string)r.Attribute("Id"))).ToList().ForEach(r => r.Remove());
                entry.Delete();
                using (var w = zip.CreateEntry(pair.Key).Open()) doc.Save(w);
            }
        }

        static bool IsVml(string name) => name.StartsWith("xl/drawings/", StringComparison.Ordinal) && name.EndsWith(".vml", StringComparison.Ordinal);

        static void FixVml(ZipArchive zip)
        {
            foreach (var entry in zip.Entries.Where(e => IsVml(e.FullName)).ToList())
            {
                string text;
                using (var r = new StreamReader(entry.Open())) text = r.ReadToEnd();
                var fixedText = ShapeTypeId.Replace(text, "$1_x0000_t202$2");
                if (fixedText == text) continue;
                var name = entry.FullName;
                entry.Delete();
                using (var w = new StreamWriter(zip.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false))) w.Write(fixedText);
            }
        }
    }
}
