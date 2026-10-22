using System;
using OpenSheet.Cells;
using OpenSheet.Internal;
using OpenSheet.Internal.Formulas;
using OpenSheet.Styles;
using XL = NPOI.SS.UserModel;

namespace OpenSheet
{
    /// <summary>One cell of a worksheet. Cells are views: they stay valid whether or not the cell holds anything yet.</summary>
    public class Cell : IStylable
    {
        internal Cell(WorkSheet sheet, int row, int column) { Sheet = sheet; RowIndex = row; ColumnIndex = column; }

        internal WorkSheet Sheet { get; }
        WorkBook Book => Sheet.Book;

        /// <summary>Zero-based. Row 3 in Excel has index 2.</summary>
        public int RowIndex { get; }
        /// <summary>Zero-based. Column "A" has index 0.</summary>
        public int ColumnIndex { get; }

        public RangeAddress Address => new RangeAddress(RowIndex, RowIndex, ColumnIndex, ColumnIndex);
        public string AddressString => A1.Cell(RowIndex, ColumnIndex);
        public string Location => AddressString;
        public IStyle Style => new RangeStyle(new Range(Sheet, Address));

        internal XL.ICell Peek() => Sheet.Native.GetRow(RowIndex)?.GetCell(ColumnIndex);

        internal XL.ICell Ensure()
        {
            var row = Sheet.Native.GetRow(RowIndex) ?? Sheet.Native.CreateRow(RowIndex);
            return row.GetCell(ColumnIndex) ?? row.CreateCell(ColumnIndex);
        }

        internal void Remove()
        {
            var c = Peek();
            if (c != null) c.Row.RemoveCell(c);
            Book.Touch();
        }

        internal object Read(out Kind kind) => CellValues.Read(Book, Peek(), out kind);

        // ---- value ----

        /// <summary>
        /// double for numbers, DateTime for date-formatted numbers, string, bool, or the error text (e.g. "#DIV/0!").
        /// Null for an empty cell. Setting null, "" or whitespace removes the cell.
        /// </summary>
        public object Value
        {
            get => Read(out _);
            set
            {
                if (CellValues.IsClearing(value)) { Remove(); return; }
                CellValues.Write(Book, Ensure(), value);
                Book.Touch();
            }
        }

        public string Text { get => StringValue; set => StringValue = value; }

        /// <summary>The raw value as text: numbers in invariant form (0.1234, not 12.34%). See <see cref="FormattedCellValue"/> for display text.</summary>
        public virtual string StringValue
        {
            get => CellValues.Raw(Read(out _));
            set => Value = value;
        }

        public virtual bool BoolValue
        {
            get
            {
                var v = Read(out var k);
                switch (k)
                {
                    case Kind.Bool: return (bool)v;
                    case Kind.Number: return (double)v > 0;
                    case Kind.Text: return bool.TryParse((string)v, out var b) && b;
                    default: return false;
                }
            }
            set => Value = value;
        }

        double Num() => CellValues.ToDouble(Read(out var k) is object v && (k == Kind.Number || k == Kind.Date) ? v : null);

        public virtual double DoubleValue { get => Num(); set => Value = value; }
        public virtual decimal DecimalValue { get => (decimal)Num(); set => Value = value; }
        public virtual float FloatValue { get => (float)Num(); set => Value = value; }
        public virtual int IntValue { get => Conv.ToInt(Num()); set => Value = value; }
        public virtual long LongValue { get => Conv.ToLong(Num()); set => Value = value; }

        [Obsolete("Use LongValue instead.")]
        public virtual long Int64Value { get => LongValue; set => LongValue = value; }

        [Obsolete("Use IntValue instead.")]
        public virtual int Int32Value { get => IntValue; set => IntValue = value; }

        public DateTime? DateTimeValue
        {
            get => Read(out var k) is DateTime d && k == Kind.Date ? d : (DateTime?)null;
            set { if (value.HasValue) Value = value.Value; else Remove(); }
        }

        public void SetDateTimeValue(DateTime value, string format)
        {
            var c = Ensure();
            CellValues.WriteDate(Book, c, value);
            FormatString = format;
        }

        public TType GetValue<TType>() => Conv.To<TType>(Value);

        public bool TryGetValue<TType>(out TType value)
        {
            try { value = GetValue<TType>(); return true; }
            catch (Exception) { value = default(TType); return false; }
        }

        public override string ToString() => StringValue;

        // ---- kind ----

        public bool IsEmpty { get { Read(out var k); return k == Kind.Empty; } }
        public bool IsNumeric { get { Read(out var k); return k == Kind.Number; } }
        public bool IsDateTime { get { Read(out var k); return k == Kind.Date; } }
        public bool IsBoolean { get { Read(out var k); return k == Kind.Bool; } }
        public bool IsText { get { Read(out var k); return k == Kind.Text; } }
        public bool IsError { get { Read(out var k); return k == Kind.Error; } }
        public bool IsFormula => Peek()?.CellType == XL.CellType.Formula;

        // ---- formatting ----

        public string FormatString
        {
            get => (Peek()?.CellStyle ?? Book.Native.GetCellStyleAt(0)).GetDataFormatString();
            set
            {
                var fmt = value ?? "General";
                var idx = Book.Native.CreateDataFormat().GetFormat(fmt);
                Book.Styles.Apply(Ensure(), "numfmt:" + fmt, s => s.DataFormat = idx);
                Book.Touch();
            }
        }

        /// <summary>The text Excel would display, formatting applied. Uses the invariant culture.</summary>
        public string FormattedCellValue
        {
            get
            {
                var c = Peek();
                if (c == null) return "";
                try { return Display.Text(Book, c); }
                catch (Exception) { return StringValue; }
            }
        }

        // ---- formula ----

        /// <summary>The formula with a leading "=", or null when the cell holds no formula. Set with or without "=".</summary>
        public string Formula
        {
            get => IsFormula ? "=" + FormulaDialect.ToDisplay(Peek().CellFormula) : null;
            set
            {
                if (string.IsNullOrWhiteSpace(value)) { RemoveFormula(); return; }
                Ensure().SetCellFormula(FormulaDialect.ToStored(value.TrimStart().TrimStart('=')));
                Book.Touch();
            }
        }

        public void RemoveFormula()
        {
            if (IsFormula) { Peek().RemoveFormula(); Book.Touch(); }
        }

        /// <summary>Computes the formula and stores the result in the file; the cell stays a formula cell.</summary>
        public Cell Evaluate()
        {
            var c = Peek();
            if (c?.CellType == XL.CellType.Formula) CellValues.IsUnsupported(Book.Evaluator, c); // an unsupported function leaves the cell as it was
            return this;
        }

        // ---- hyperlink / comment ----

        /// <summary>
        /// A URL (http, https, ftp), "mailto:" address, "file:///" path, or a cell/defined name inside the workbook.
        /// </summary>
        public string Hyperlink
        {
            get => Peek()?.Hyperlink?.Address;
            set
            {
                if (string.IsNullOrWhiteSpace(value)) { RemoveHyperlink(); return; }
                var type = XL.HyperlinkType.Document;
                if (value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) type = XL.HyperlinkType.Email;
                else if (value.StartsWith("file:", StringComparison.OrdinalIgnoreCase)) type = XL.HyperlinkType.File;
                else if (value.IndexOf("://", StringComparison.Ordinal) > 0) type = XL.HyperlinkType.Url;
                var link = Book.Native.GetCreationHelper().CreateHyperlink(type);
                link.Address = value;
                Peek()?.RemoveHyperlink(); // NPOI appends; the getter would keep returning the first link ever set
                Ensure().Hyperlink = link;
                Book.Touch();
            }
        }

        public void RemoveHyperlink() { Peek()?.RemoveHyperlink(); }

        /// <summary>The comment on this cell, or null.</summary>
        public Comment Comment
        {
            get { var cm = Peek()?.CellComment; return cm == null ? null : new Comment(cm, Book.Native.GetCreationHelper()); }
        }

        public Comment AddComment(string content, string author, bool isVisible = false)
        {
            var cell = Ensure();
            var helper = Book.Native.GetCreationHelper();
            var anchor = helper.CreateClientAnchor();
            anchor.Col1 = ColumnIndex + 1; anchor.Col2 = ColumnIndex + 3;
            anchor.Row1 = RowIndex; anchor.Row2 = RowIndex + 3;
            var cm = Sheet.Native.CreateDrawingPatriarch().CreateCellComment(anchor);
            cm.String = helper.CreateRichTextString(content ?? "");
            cm.Author = author;
            cm.Visible = isVisible;
            cell.CellComment = cm;
            Book.Touch();
            return new Comment(cm, helper);
        }

        public void RemoveComment() { Peek()?.RemoveCellComment(); }

        // ---- copy ----

        /// <summary>Copies value, formula and style to <paramref name="pointToInsert"/> (e.g. "C4") on <paramref name="sheet"/>.</summary>
        public void Copy(WorkSheet sheet, string pointToInsert)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            A1.ParseCell(pointToInsert, out var r, out var c);
            Sheet.CopyCellTo(RowIndex, ColumnIndex, sheet, r, c);
        }
    }
}
