using System;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal;

namespace OpenSheet.Metadata
{
    /// <summary>Named values stored with the document, beyond the standard fields.</summary>
    public interface ICustomProperties
    {
        /// <summary>Adds or replaces a property. Supported: string, int, double, bool; other types are stored as text. XLS also keeps DateTime.</summary>
        void AddProperty(string name, object value);
        bool Contains(string name);
        object GetValue(string name);
        T GetValue<T>(string name);
        bool TryGetValue<T>(string name, out T value);
        bool Remove(string name);
    }

    /// <summary>Document properties (File → Info in Excel).</summary>
    public interface IMetadata
    {
        string Author { get; set; }
        string Title { get; set; }
        string Subject { get; set; }
        string ApplicationName { get; }
        string Category { get; set; }
        string Company { get; }
        string Manager { get; }
        string Template { get; }
        string Comments { get; set; }
        string Keywords { get; set; }
        DateTime? Created { get; set; }
        DateTime? Modified { get; set; }
        DateTime? LastPrinted { get; set; }
        ICustomProperties CustomProperties { get; }
    }

    abstract class CustomPropertiesBase : ICustomProperties
    {
        /// <summary>Every property name, for copying a workbook to the other file format.</summary>
        internal abstract System.Collections.Generic.IEnumerable<string> Names { get; }

        public abstract void AddProperty(string name, object value);
        public abstract bool Contains(string name);
        public abstract object GetValue(string name);
        public abstract bool Remove(string name);

        public T GetValue<T>(string name)
        {
            if (!Contains(name)) throw new ArgumentException($"No custom property named '{name}'.", nameof(name));
            return Conv.To<T>(GetValue(name));
        }

        public bool TryGetValue<T>(string name, out T value)
        {
            try { value = GetValue<T>(name); return true; }
            catch (Exception) { value = default(T); return false; }
        }

        protected static void Require(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("The property name is empty.", nameof(name));
        }
    }

    sealed class XlsxCustomProperties : CustomPropertiesBase
    {
        readonly NPOI.CustomProperties _p;
        public XlsxCustomProperties(NPOI.CustomProperties p) { _p = p; }

        public override void AddProperty(string name, object value)
        {
            Require(name);
            Remove(name);
            switch (value)
            {
                case int i: _p.AddProperty(name, i); break;
                case short _: case byte _: case sbyte _: case ushort _: _p.AddProperty(name, Convert.ToInt32(value)); break;
                case long _: case uint _: case ulong _: case float _: case double _: case decimal _:
                    _p.AddProperty(name, Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture)); break;
                case bool b: _p.AddProperty(name, b); break;
                case DateTime d: _p.AddProperty(name, d.ToString("o", System.Globalization.CultureInfo.InvariantCulture)); break; // shortcut: stored as ISO text, no native date type in NPOI's OOXML API
                default: _p.AddProperty(name, Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? ""); break;
            }
        }

        internal override System.Collections.Generic.IEnumerable<string> Names => _p.GetUnderlyingProperties().property.Select(x => x.name).ToList();
        public override bool Contains(string name) => _p.Contains(name);
        public override object GetValue(string name) => _p.GetProperty(name)?.Item;
        public override bool Remove(string name) => _p.GetUnderlyingProperties().property.RemoveAll(x => x.name == name) > 0;
    }

    sealed class XlsxMetadata : IMetadata
    {
        readonly NPOI.POIXMLProperties _p;
        public XlsxMetadata(XSSFWorkbook wb) { _p = wb.GetProperties(); }

        NPOI.CoreProperties Core => _p.CoreProperties;
        public string Author { get => Core.Creator; set => Core.Creator = value; }
        public string Title { get => Core.Title; set => Core.Title = value; }
        public string Subject { get => Core.Subject; set => Core.Subject = value; }
        public string Category { get => Core.Category; set => Core.Category = value; }
        public string Comments { get => Core.Description; set => Core.Description = value; }
        public string Keywords { get => Core.Keywords; set => Core.Keywords = value; }
        public DateTime? Created { get => Core.Created; set => Core.Created = value; }
        public DateTime? Modified { get => Core.Modified; set => Core.Modified = value; }
        public DateTime? LastPrinted { get => Core.LastPrinted; set => Core.LastPrinted = value; }
        public string ApplicationName => _p.ExtendedProperties.Application;
        public string Company => _p.ExtendedProperties.Company;
        public string Manager => _p.ExtendedProperties.Manager;
        public string Template => _p.ExtendedProperties.Template;
        public ICustomProperties CustomProperties => new XlsxCustomProperties(_p.CustomProperties);
    }

    sealed class XlsCustomProperties : CustomPropertiesBase
    {
        readonly HSSFWorkbook _wb;
        public XlsCustomProperties(HSSFWorkbook wb) { _wb = wb; }

        NPOI.HPSF.CustomProperties Props => _wb.DocumentSummaryInformation.CustomProperties ?? new NPOI.HPSF.CustomProperties();

        // HPSF's own Remove does not stick after a reload, so the set is rebuilt from scratch on every change.
        void Rewrite(Action<System.Collections.Generic.Dictionary<string, object>> change)
        {
            var old = Props;
            var all = new System.Collections.Generic.Dictionary<string, object>();
            foreach (var n in old.NameSet()) all[n] = old.Get(n);
            change(all);
            var d = _wb.DocumentSummaryInformation;
            if (d.CustomProperties != null) d.RemoveCustomProperties(); // throws when the second section does not exist yet
            var fresh = new NPOI.HPSF.CustomProperties();
            foreach (var kv in all) fresh.Put(kv.Key, kv.Value);
            d.CustomProperties = fresh;
        }

        public override void AddProperty(string name, object value)
        {
            Require(name);
            object v;
            switch (value)
            {
                case string _: case bool _: case DateTime _: case int _: case long _: case double _: v = value; break;
                case short _: case byte _: case sbyte _: case ushort _: v = Convert.ToInt32(value); break;
                case float _: case decimal _: case uint _: case ulong _: v = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture); break;
                default: v = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? ""; break;
            }
            Rewrite(all => all[name] = v);
        }

        internal override System.Collections.Generic.IEnumerable<string> Names => Props.NameSet().ToList();
        public override bool Contains(string name) => Props.ContainsKey(name);
        public override object GetValue(string name) => Props.Get(name);

        public override bool Remove(string name)
        {
            if (!Contains(name)) return false;
            Rewrite(all => all.Remove(name));
            return true;
        }
    }

    sealed class XlsMetadata : IMetadata
    {
        readonly HSSFWorkbook _wb;

        public XlsMetadata(HSSFWorkbook wb)
        {
            _wb = wb;
            if (wb.SummaryInformation == null || wb.DocumentSummaryInformation == null) wb.CreateInformationProperties();
        }

        SummaryInformation S => _wb.SummaryInformation;
        DocumentSummaryInformation D => _wb.DocumentSummaryInformation;

        public string Author { get => S.Author; set => S.Author = value; }
        public string Title { get => S.Title; set => S.Title = value; }
        public string Subject { get => S.Subject; set => S.Subject = value; }
        public string Comments { get => S.Comments; set => S.Comments = value; }
        public string Keywords { get => S.Keywords; set => S.Keywords = value; }
        public string Category { get => D.Category; set => D.Category = value; }
        public DateTime? Created { get => S.CreateDateTime; set => S.CreateDateTime = value; }
        public DateTime? Modified { get => S.LastSaveDateTime; set => S.LastSaveDateTime = value; }
        public DateTime? LastPrinted { get => S.LastPrinted; set => S.LastPrinted = value; }
        public string ApplicationName => S.ApplicationName;
        public string Company => D.Company;
        public string Manager => D.Manager;
        public string Template => S.Template;
        public ICustomProperties CustomProperties => new XlsCustomProperties(_wb);
    }

    static class MetadataCopier
    {
        /// <summary>Copies the document fields that can be set, and the custom properties.</summary>
        public static void Copy(IMetadata from, IMetadata to)
        {
            // HPSF refuses a null string, so only fields that hold something are copied
            if (from.Author != null) to.Author = from.Author;
            if (from.Title != null) to.Title = from.Title;
            if (from.Subject != null) to.Subject = from.Subject;
            if (from.Category != null) to.Category = from.Category;
            if (from.Comments != null) to.Comments = from.Comments;
            if (from.Keywords != null) to.Keywords = from.Keywords;
            if (from.Created != null) to.Created = from.Created;
            if (from.Modified != null) to.Modified = from.Modified;
            if (from.LastPrinted != null) to.LastPrinted = from.LastPrinted;
            if (from.CustomProperties is CustomPropertiesBase props)
                foreach (var name in props.Names) to.CustomProperties.AddProperty(name, from.CustomProperties.GetValue(name));
        }
    }
}
