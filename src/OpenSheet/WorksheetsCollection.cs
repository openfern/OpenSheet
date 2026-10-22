using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OpenSheet
{
    /// <summary>The worksheets of a workbook, in tab order. A live view: changes apply to the workbook.</summary>
    public class WorksheetsCollection : IList<WorkSheet>
    {
        readonly WorkBook _book;
        internal WorksheetsCollection(WorkBook book) { _book = book; }

        public int Count => _book.Native.NumberOfSheets;
        public bool IsReadOnly => false;

        public WorkSheet this[int index]
        {
            get
            {
                if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
                return _book.Wrap(_book.Native.GetSheetAt(index));
            }
            set
            {
                if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
                if (value == null) throw new ArgumentNullException(nameof(value));
                if (ReferenceEquals(value.Book, _book)) throw new ArgumentException("The worksheet already belongs to this workbook.", nameof(value));
                _book.RemoveWorkSheet(index);   // a copy of the new sheet takes its place; formulas elsewhere that pointed at the old one become #REF!, as in Excel
                Insert(index, value);
            }
        }

        public IEnumerator<WorkSheet> GetEnumerator() => _book.SheetList().GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public WorkSheet Create(string name) => _book.CreateWorkSheet(name);

        /// <summary>Appends a copy of a sheet from another workbook. A sheet that already belongs to this workbook is rejected.</summary>
        public void Add(WorkSheet item)
        {
            if (item == null) throw new ArgumentNullException(nameof(item));
            if (ReferenceEquals(item.Book, _book)) throw new ArgumentException("The worksheet already belongs to this workbook.", nameof(item));
            item.CopyTo(_book, item.Name);
        }

        public void Insert(int index, WorkSheet item)
        {
            Add(item);
            _book.Native.SetSheetOrder(item.Name, index);
        }

        public void Clear()
        {
            for (var i = Count - 1; i >= 0; i--) _book.RemoveWorkSheet(i);
        }

        public bool Contains(WorkSheet item) => item != null && ReferenceEquals(item.Book, _book) && IndexOf(item) >= 0;

        public void CopyTo(WorkSheet[] array, int arrayIndex) => _book.SheetList().CopyTo(array, arrayIndex);

        public bool Remove(WorkSheet item)
        {
            var i = IndexOf(item);
            if (i < 0) return false;
            RemoveAt(i);
            return true;
        }

        public void RemoveAt(int index) => _book.RemoveWorkSheet(index);

        public int FindIndex(Predicate<WorkSheet> predicate) => _book.SheetList().FindIndex(predicate);

        public int IndexOf(WorkSheet item) => item == null || !ReferenceEquals(item.Book, _book) ? -1 : _book.Native.GetSheetIndex(item.Native);

        public void SetSheetIndex(WorkSheet sheet, int index) => _book.SetSheetPosition(sheet.Name, index);
    }
}
