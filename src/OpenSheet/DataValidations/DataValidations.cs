using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NPOI.SS.Util;
using OpenSheet.Internal;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.DataValidations
{
    /// <summary>How the entered value is compared with the rule's limits. The numbers match Excel's file values.</summary>
    public enum ComparisonOperator
    {
        None = -1, Between = 0, NotBetween = 1, Equal = 2, NotEqual = 3, GreaterThan = 4, LessThan = 5, GreaterThanOrEqual = 6, LessThanOrEqual = 7,
    }

    public enum ConstraintType { AnyData = 0, WholeNumber = 1, Decimal = 2, List = 3, Date = 4, Time = 5, TextLength = 6, Custom = 7 }

    /// <summary>Icon shown in the error alert when the rule is broken.</summary>
    public enum ErrorAlert { Stop = 0, Warning = 1, Information = 2 }

    /// <summary>A rule that limits what can be typed into a range of cells.</summary>
    public class DataValidation
    {
        readonly NPOI.OpenXmlFormats.Spreadsheet.CT_DataValidation _ct; // XLSX only: NPOI's own flag setter does not reach the file
        readonly NPOI.HSSF.UserModel.HSSFSheet _xls;                    // XLS only: validations are rebuilt from records, edits must be written back
        int _xlsIndex;

        internal DataValidation(XL.IDataValidation native, NPOI.OpenXmlFormats.Spreadsheet.CT_DataValidation ct, NPOI.HSSF.UserModel.HSSFSheet xls, int index)
        {
            Native = native; _ct = ct; _xls = xls; _xlsIndex = index;
        }

        internal XL.IDataValidation Native { get; }

        // XLS: replace the stored record with one built from the edited object (it moves to the end of the list).
        void Changed()
        {
            if (_xls == null) return;
            HssfValidations.RemoveAt(_xls, _xlsIndex);
            _xls.AddValidationData(Native);
            _xlsIndex = _xls.GetDataValidations().Count - 1;
        }

        string Limit(string text, bool second)
        {
            if (text != null) return text;
            // XLS keeps numeric limits as numbers, not formula text
            if (Native.ValidationConstraint is NPOI.HSSF.UserModel.DVConstraint d && ConstraintType != ConstraintType.List && ConstraintType != ConstraintType.Custom)
            {
                var v = second ? d.Value2 : d.Value1;
                if (!double.IsNaN(v)) return v.ToString("R", CultureInfo.InvariantCulture);
            }
            return null;
        }

        /// <summary>The first range the rule covers (a rule loaded from a file may list several).</summary>
        public RangeAddress ConstrainedRangeAddress => RangeAddress.From(Native.Regions.CellRangeAddresses[0]);

        public ConstraintType ConstraintType => (ConstraintType)Native.ValidationConstraint.GetValidationType();

        /// <summary>Only meaningful for whole number, decimal, date, time and text length rules; None otherwise.</summary>
        public ComparisonOperator ComparisonOperator
        {
            get
            {
                switch (ConstraintType)
                {
                    case ConstraintType.WholeNumber: case ConstraintType.Decimal: case ConstraintType.Date: case ConstraintType.Time: case ConstraintType.TextLength:
                        return (ComparisonOperator)Native.ValidationConstraint.Operator;
                    default: return ComparisonOperator.None;
                }
            }
            set
            {
                Native.ValidationConstraint.Operator = (int)value;
                if (_ct != null && value != ComparisonOperator.None) _ct.@operator = XmlOperators[(int)value];
                Changed();
            }
        }

        // index = our (and NPOI's) operator number, value = the attribute Excel stores
        static readonly NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator[] XmlOperators =
        {
            NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.between, NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.notBetween,
            NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.equal, NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.notEqual,
            NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.greaterThan, NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.lessThan,
            NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.greaterThanOrEqual, NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationOperator.lessThanOrEqual,
        };

        public ErrorAlert ErrorAlert { get => (ErrorAlert)Native.ErrorStyle; set { Native.ErrorStyle = (int)value; Changed(); } }

        /// <summary>For a list rule: the list (as typed in Excel). For a custom rule: the formula.</summary>
        public string Formula1 => Limit(Native.ValidationConstraint.Formula1, false);
        /// <summary>The upper limit of a "between" style rule; otherwise null.</summary>
        public string Formula2 =>
            ComparisonOperator == ComparisonOperator.Between || ComparisonOperator == ComparisonOperator.NotBetween
                ? Limit(Native.ValidationConstraint.Formula2, true) : Native.ValidationConstraint.Formula2;

        public bool EmptyCellAllowed { get => Native.EmptyCellAllowed; set { Native.EmptyCellAllowed = value; Changed(); } }
        public bool ShowErrorBox { get => Native.ShowErrorBox; set { Native.ShowErrorBox = value; Changed(); } }
        public bool ShowPromptBox { get => Native.ShowPromptBox; set { Native.ShowPromptBox = value; Changed(); } }

        /// <summary>The in-cell drop-down arrow; only list rules have one (false for the rest).</summary>
        public bool ShowDropDownList
        {
            // In the XLSX file the attribute is "showDropDown", and true means the arrow is HIDDEN.
            get => ConstraintType == ConstraintType.List && (_ct != null ? !_ct.showDropDown : !Native.SuppressDropDownArrow);
            set { if (_ct != null) _ct.showDropDown = !value; else { Native.SuppressDropDownArrow = !value; Changed(); } }
        }

        public string ErrorBoxTitle { get => Native.ErrorBoxTitle; set { Native.CreateErrorBox(value, Native.ErrorBoxText); Changed(); } }
        public string ErrorBoxText { get => Native.ErrorBoxText; set { Native.CreateErrorBox(Native.ErrorBoxTitle, value); Changed(); } }
        public string PromptBoxTitle { get => Native.PromptBoxTitle; set { Native.CreatePromptBox(value, Native.PromptBoxText); Changed(); } }
        public string PromptBoxText { get => Native.PromptBoxText; set { Native.CreatePromptBox(Native.PromptBoxTitle, value); Changed(); } }

        public override bool Equals(object obj) =>
            obj is DataValidation o && ConstrainedRangeAddress.IsEqual(o.ConstrainedRangeAddress) && ConstraintType == o.ConstraintType
            && ComparisonOperator == o.ComparisonOperator && Formula1 == o.Formula1 && Formula2 == o.Formula2;

        public override int GetHashCode() => ConstrainedRangeAddress.GetHashCode() ^ ((int)ConstraintType << 24) ^ (Formula1?.GetHashCode() ?? 0);
    }

    /// <summary>The data validation rules of one worksheet (a live view), or a filtered selection of them.</summary>
    public class DataValidationsCollection : IEnumerable<DataValidation>
    {
        readonly WorkSheet _sheet;
        readonly Func<DataValidation, bool> _filter;

        internal DataValidationsCollection(WorkSheet sheet, Func<DataValidation, bool> filter = null) { _sheet = sheet; _filter = filter; }

        // NPOI builds the wrappers from the sheet's validation elements in order, so index i is element i.
        List<DataValidation> Items()
        {
            var natives = _sheet.Native.GetDataValidations() ?? new List<XL.IDataValidation>();
            var cts = (_sheet.Native as NPOI.XSSF.UserModel.XSSFSheet)?.GetCTWorksheet().dataValidations?.dataValidation;
            var xls = _sheet.Native as NPOI.HSSF.UserModel.HSSFSheet;
            return natives.Select((v, i) => new DataValidation(v, cts != null && i < cts.Count ? cts[i] : null, xls, i))
                .Where(v => _filter == null || _filter(v)).ToList();
        }

        public DataValidation this[int index] => Items()[index];
        public int Count => Items().Count;
        public IEnumerator<DataValidation> GetEnumerator() => Items().GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        static bool Touches(DataValidation v, RangeAddress a) => v.Native.Regions.CellRangeAddresses.Any(r => RangeAddress.From(r).GetIntersection(a) != null);

        public DataValidation GetRuleForCell(Cell constrainedCell) =>
            Items().FirstOrDefault(v => Touches(v, constrainedCell.Address));

        public DataValidationsCollection GetRulesForRange(string constrainedAddress) => GetRulesForRange(A1.Parse(constrainedAddress));
        public DataValidationsCollection GetRulesForRange(Range constrainedRange) => GetRulesForRange(constrainedRange.RangeAddress);
        public DataValidationsCollection GetRulesForRange(RangeAddress rangeAddress) => new DataValidationsCollection(_sheet, v => Touches(v, rangeAddress));

        public bool RangeAlreadyHasDataValidation(RangeAddress rangeAddress) => Items().Any(v => Touches(v, rangeAddress));

        // ---- adding ----

        DataValidation Add(RangeAddress a, XL.IDataValidationConstraint constraint)
        {
            var native = _sheet.Native.GetDataValidationHelper().CreateValidation(constraint, new CellRangeAddressList(a.FirstRow, a.LastRow, a.FirstColumn, a.LastColumn));
            native.ShowErrorBox = true;
            native.EmptyCellAllowed = true;
            _sheet.Native.AddValidationData(native);
            return new DataValidationsCollection(_sheet).Items().Last(); // re-read so the wrapper is bound to the stored element
        }

        XL.IDataValidationHelper Helper => _sheet.Native.GetDataValidationHelper();

        static void NeedsMax(ComparisonOperator comparisonOperator, bool hasMax)
        {
            if ((comparisonOperator == ComparisonOperator.Between || comparisonOperator == ComparisonOperator.NotBetween) && !hasMax)
                throw new ArgumentException("A between / not-between rule needs both limits.");
            if (comparisonOperator == ComparisonOperator.None) throw new ArgumentException("Choose a comparison operator.");
        }

        static string Num(decimal d) => d.ToString(CultureInfo.InvariantCulture);

        public DataValidation AddFormulaListRule(string constrainedAddress, string formula) => AddFormulaListRule(A1.Parse(constrainedAddress), formula);
        public DataValidation AddFormulaListRule(Cell constrainedCell, string formula) => AddFormulaListRule(constrainedCell.Address, formula);
        public DataValidation AddFormulaListRule(Range constrainedRange, string formula) => AddFormulaListRule(constrainedRange.RangeAddress, formula);
        public DataValidation AddFormulaListRule(RangeAddress constrainedRangeAddress, string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("The list formula is empty.", nameof(formula));
            return Add(constrainedRangeAddress, Helper.CreateFormulaListConstraint(formula.TrimStart('=')));
        }

        public DataValidation AddStringListRule(string constrainedAddress, string[] values) => AddStringListRule(A1.Parse(constrainedAddress), values);
        public DataValidation AddStringListRule(Cell constrainedCell, string[] values) => AddStringListRule(constrainedCell.Address, values);
        public DataValidation AddStringListRule(Range constrainedRange, string[] values) => AddStringListRule(constrainedRange.RangeAddress, values);
        public DataValidation AddStringListRule(RangeAddress constrainedRangeAddress, string[] values)
        {
            if (values == null || values.Length == 0) throw new ArgumentException("Give at least one value.", nameof(values));
            if (string.Join(",", values).Length > 255) throw new ArgumentException("Excel limits an inline list to 255 characters; use AddFormulaListRule with a cell range instead.", nameof(values));
            return Add(constrainedRangeAddress, Helper.CreateExplicitListConstraint(values));
        }

        public DataValidation AddIntegerRule(string constrainedAddress, ComparisonOperator comparisonOperator, int minValue, int? maxValue) => AddIntegerRule(A1.Parse(constrainedAddress), comparisonOperator, minValue, maxValue);
        public DataValidation AddIntegerRule(Cell constrainedCell, ComparisonOperator comparisonOperator, int minValue, int? maxValue) => AddIntegerRule(constrainedCell.Address, comparisonOperator, minValue, maxValue);
        public DataValidation AddIntegerRule(Range constrainedRange, ComparisonOperator comparisonOperator, int minValue, int? maxValue) => AddIntegerRule(constrainedRange.RangeAddress, comparisonOperator, minValue, maxValue);
        public DataValidation AddIntegerRule(RangeAddress constrainedRangeAddress, ComparisonOperator comparisonOperator, int minValue, int? maxValue)
        {
            NeedsMax(comparisonOperator, maxValue.HasValue);
            return Add(constrainedRangeAddress, Helper.CreateintConstraint((int)comparisonOperator, minValue.ToString(CultureInfo.InvariantCulture), maxValue?.ToString(CultureInfo.InvariantCulture)));
        }

        public DataValidation AddDecimalRule(string constrainedAddress, ComparisonOperator comparisonOperator, decimal minValue, decimal? maxValue) => AddDecimalRule(A1.Parse(constrainedAddress), comparisonOperator, minValue, maxValue);
        public DataValidation AddDecimalRule(Cell constrainedCell, ComparisonOperator comparisonOperator, decimal minValue, decimal? maxValue) => AddDecimalRule(constrainedCell.Address, comparisonOperator, minValue, maxValue);
        public DataValidation AddDecimalRule(Range constrainedRange, ComparisonOperator comparisonOperator, decimal minValue, decimal? maxValue) => AddDecimalRule(constrainedRange.RangeAddress, comparisonOperator, minValue, maxValue);
        public DataValidation AddDecimalRule(RangeAddress constrainedRangeAddress, ComparisonOperator comparisonOperator, decimal minValue, decimal? maxValue)
        {
            NeedsMax(comparisonOperator, maxValue.HasValue);
            return Add(constrainedRangeAddress, Helper.CreateDecimalConstraint((int)comparisonOperator, Num(minValue), maxValue.HasValue ? Num(maxValue.Value) : null));
        }

        public DataValidation AddDateRule(string constrainedAddress, ComparisonOperator comparisonOperator, DateTime minValue, DateTime? maxValue) => AddDateRule(A1.Parse(constrainedAddress), comparisonOperator, minValue, maxValue);
        public DataValidation AddDateRule(Cell constrainedCell, ComparisonOperator comparisonOperator, DateTime minValue, DateTime? maxValue) => AddDateRule(constrainedCell.Address, comparisonOperator, minValue, maxValue);
        public DataValidation AddDateRule(Range constrainedRange, ComparisonOperator comparisonOperator, DateTime minValue, DateTime? maxValue) => AddDateRule(constrainedRange.RangeAddress, comparisonOperator, minValue, maxValue);
        public DataValidation AddDateRule(RangeAddress constrainedRangeAddress, ComparisonOperator comparisonOperator, DateTime minValue, DateTime? maxValue)
        {
            NeedsMax(comparisonOperator, maxValue.HasValue);
            // NPOI turns "yyyy-MM-dd" text into the Excel date serial; a time of day is dropped.
            const string fmt = "yyyy-MM-dd";
            string Text(DateTime d) => d.ToString(fmt, CultureInfo.InvariantCulture);
            return Add(constrainedRangeAddress, Helper.CreateDateConstraint((int)comparisonOperator, Text(minValue), maxValue.HasValue ? Text(maxValue.Value) : null, fmt));
        }

        public DataValidation AddTextLengthRule(string constrainedAddress, ComparisonOperator comparisonOperator, int minLength, int? maxLength) => AddTextLengthRule(A1.Parse(constrainedAddress), comparisonOperator, minLength, maxLength);
        public DataValidation AddTextLengthRule(Cell constrainedCell, ComparisonOperator comparisonOperator, int minLength, int? maxLength) => AddTextLengthRule(constrainedCell.Address, comparisonOperator, minLength, maxLength);
        public DataValidation AddTextLengthRule(Range constrainedRange, ComparisonOperator comparisonOperator, int minLength, int? maxLength) => AddTextLengthRule(constrainedRange.RangeAddress, comparisonOperator, minLength, maxLength);
        public DataValidation AddTextLengthRule(RangeAddress constrainedRangeAddress, ComparisonOperator comparisonOperator, int minLength, int? maxLength)
        {
            NeedsMax(comparisonOperator, maxLength.HasValue);
            return Add(constrainedRangeAddress, Helper.CreateTextLengthConstraint((int)comparisonOperator, minLength.ToString(CultureInfo.InvariantCulture), maxLength?.ToString(CultureInfo.InvariantCulture)));
        }

        // ---- removing ----

        public void Remove(DataValidation dataValidation)
        {
            if (dataValidation == null) throw new ArgumentNullException(nameof(dataValidation));
            if (_sheet.Native is NPOI.HSSF.UserModel.HSSFSheet xls)
            {
                // NPOI ignores RemoveDataValidation for XLS; remove the stored record itself.
                var i = new DataValidationsCollection(_sheet).Items().FindIndex(x => x.Equals(dataValidation));
                if (i >= 0) HssfValidations.RemoveAt(xls, i);
            }
            else _sheet.Native.RemoveDataValidation(dataValidation.Native);
        }

        public void RemoveAt(int index) => Remove(Items()[index]);

        /// <summary>Removes every rule that overlaps the range (a rule is never split).</summary>
        public void RemoveAllRulesFromRange(string address) => RemoveAllRulesFromRange(A1.Parse(address));
        public void RemoveAllRulesFromRange(Range constrainedRange) => RemoveAllRulesFromRange(constrainedRange.RangeAddress);
        public void RemoveAllRulesFromRange(RangeAddress rangeAddress)
        {
            foreach (var v in Items().Where(v => Touches(v, rangeAddress))) Remove(v);
        }

        public void RemoveRuleFromCell(Cell constrainedCell) => RemoveAllRulesFromRange(constrainedCell.Address);
    }
}
