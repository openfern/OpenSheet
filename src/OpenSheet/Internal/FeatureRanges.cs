using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OpenSheet.Internal.Formulas;
using XL = NPOI.SS.UserModel;

namespace OpenSheet.Internal
{
    /// <summary>
    /// Keeps the ranges of data validations and conditional formats on the cells they cover when rows or columns are inserted or
    /// removed. NPOI moves the cells but leaves these ranges where they were.
    /// </summary>
    internal static class FeatureRanges
    {
        /// <summary>Rows (or columns) were inserted before <paramref name="at"/> (delta &gt; 0) or <c>-delta</c> removed from <paramref name="at"/>.</summary>
        public static void Adjust(WorkSheet sheet, bool rows, int at, int delta, bool includeConditionalFormats)
        {
            if (sheet.Native is XSSFSheet x)
            {
                AdjustXlsx(x, rows, at, delta, includeConditionalFormats, rows ? sheet.MaxRow : sheet.MaxColumn);
                AdjustTables(sheet, x, rows, at, delta);
            }
            else AdjustXls(sheet, rows, at, delta, includeConditionalFormats, rows ? sheet.MaxRow : sheet.MaxColumn);
        }

        // "A1", "A1:B2" or several of them separated by blanks, as stored in sqref attributes.
        static List<string> AdjustList(IEnumerable<string> refs, bool rows, int at, int delta, int max)
        {
            var result = new List<string>();
            foreach (var entry in refs ?? Enumerable.Empty<string>())
                foreach (var part in (entry ?? "").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var a = A1.Parse(part);
                    int first = rows ? a.FirstRow : a.FirstColumn, last = rows ? a.LastRow : a.LastColumn;
                    if (!References.AdjustBlock(ref first, ref last, at, delta, max)) continue;
                    var moved = rows ? new RangeAddress(first, last, a.FirstColumn, a.LastColumn) : new RangeAddress(a.FirstRow, a.LastRow, first, last);
                    result.Add(moved.Location);
                }
            return result;
        }

        static void AdjustXlsx(XSSFSheet sheet, bool rows, int at, int delta, bool conditional, int max)
        {
            var ws = sheet.GetCTWorksheet();
            var validations = ws.dataValidations;
            if (validations?.dataValidation != null)
            {
                foreach (var dv in validations.dataValidation.ToList())
                {
                    var moved = AdjustList(new[] { dv.sqref }, rows, at, delta, max);
                    if (moved.Count == 0) validations.dataValidation.Remove(dv); else dv.sqref = string.Join(" ", moved);
                }
                validations.count = (uint)validations.dataValidation.Count;
                if (validations.dataValidation.Count == 0) ws.dataValidations = null;
            }
            if (conditional && ws.conditionalFormatting != null)
            {
                foreach (var cf in ws.conditionalFormatting.ToList())
                {
                    var moved = AdjustList(new[] { cf.sqref }, rows, at, delta, max);
                    if (moved.Count == 0) ws.conditionalFormatting.Remove(cf); else cf.sqref = string.Join(" ", moved);
                }
            }
        }

        // A table moves, grows or shrinks with the rows and columns around it, and gains or loses columns when columns are
        // inserted or removed inside it (a new column is named "Column1", "Column2"... and its header cell is filled in, as Excel does).
        static void AdjustTables(WorkSheet sheet, XSSFSheet x, bool rows, int at, int delta)
        {
            var max = rows ? sheet.MaxRow : sheet.MaxColumn;
            foreach (var table in x.GetTables().ToList())
            {
                int r1 = table.StartRowIndex, r2 = table.EndRowIndex, c1 = table.StartColIndex, c2 = table.EndColIndex;
                int oldC1 = c1, oldC2 = c2;
                var ok = rows ? References.AdjustBlock(ref r1, ref r2, at, delta, max) : References.AdjustBlock(ref c1, ref c2, at, delta, max);
                if (!ok) { x.RemoveTable(table); continue; }

                var ct = table.GetCTTable();
                var hasHeader = ct.headerRowCount != 0;
                if (!rows && ct.tableColumns != null)
                {
                    var columns = ct.tableColumns.tableColumn;
                    if (delta > 0 && at > oldC1 && at <= oldC2)
                    {
                        var used = new HashSet<string>(columns.Select(c => c.name), StringComparer.OrdinalIgnoreCase);
                        var nextId = columns.Count == 0 ? 1u : columns.Max(c => c.id) + 1;
                        for (var k = 0; k < delta; k++)
                        {
                            var n = 1; while (used.Contains("Column" + n)) n++;
                            var name = "Column" + n; used.Add(name);
                            columns.Insert(at - oldC1 + k, new NPOI.OpenXmlFormats.Spreadsheet.CT_TableColumn { id = nextId++, name = name });
                            if (hasHeader) sheet.GetCellAt(r1, at + k).Value = name;
                        }
                    }
                    else if (delta < 0)
                    {
                        int from = Math.Max(at, oldC1) - oldC1, to = Math.Min(at - delta - 1, oldC2) - oldC1;
                        if (to >= from && from < columns.Count) columns.RemoveRange(from, Math.Min(to, columns.Count - 1) - from + 1);
                    }
                    ct.tableColumns.count = (uint)columns.Count;
                }

                table.CellReferences = new AreaReference(new CellReference(r1, c1), new CellReference(r2, c2));
                if (ct.autoFilter != null) ct.autoFilter.@ref = new RangeAddress(r1, hasHeader ? r2 : r1, c1, c2).Location;
            }
        }

        // XLS has no direct access to the stored ranges: each item is rebuilt with its new ranges (best effort).
        static void AdjustXls(WorkSheet sheet, bool rows, int at, int delta, bool conditional, int max)
        {
            try
            {
                var native = sheet.Native;
                var helper = native.GetDataValidationHelper();
                var xls = native as NPOI.HSSF.UserModel.HSSFSheet;
                var current = (native.GetDataValidations() ?? new List<XL.IDataValidation>()).ToList();
                for (var i = current.Count - 1; i >= 0; i--)
                {
                    var dv = current[i];
                    var before = dv.Regions.CellRangeAddresses.Select(r => RangeAddress.From(r).Location).ToList();
                    var after = AdjustList(before, rows, at, delta, max);
                    if (after.SequenceEqual(before)) continue;
                    HssfValidations.RemoveAt(xls, i);
                    if (after.Count == 0) continue;
                    var list = new CellRangeAddressList();
                    foreach (var text in after) { var a = A1.Parse(text); list.AddCellRangeAddress(a.FirstRow, a.FirstColumn, a.LastRow, a.LastColumn); }
                    var copy = helper.CreateValidation(dv.ValidationConstraint, list);
                    copy.EmptyCellAllowed = dv.EmptyCellAllowed; copy.ShowErrorBox = dv.ShowErrorBox; copy.ShowPromptBox = dv.ShowPromptBox;
                    copy.ErrorStyle = dv.ErrorStyle; copy.SuppressDropDownArrow = dv.SuppressDropDownArrow;
                    if (dv.ErrorBoxText != null || dv.ErrorBoxTitle != null) copy.CreateErrorBox(dv.ErrorBoxTitle, dv.ErrorBoxText);
                    if (dv.PromptBoxText != null || dv.PromptBoxTitle != null) copy.CreatePromptBox(dv.PromptBoxTitle, dv.PromptBoxText);
                    native.AddValidationData(copy);
                }

                if (!conditional) return;
                var cfs = native.SheetConditionalFormatting;
                for (var i = cfs.NumConditionalFormattings - 1; i >= 0; i--)
                {
                    var cf = cfs.GetConditionalFormattingAt(i);
                    var before = cf.GetFormattingRanges().Select(r => RangeAddress.From(r).Location).ToList();
                    var after = AdjustList(before, rows, at, delta, max);
                    if (after.SequenceEqual(before)) continue;
                    var rules = Enumerable.Range(0, cf.NumberOfRules).Select(cf.GetRule).ToArray();
                    cfs.RemoveConditionalFormatting(i);
                    if (after.Count > 0) cfs.AddConditionalFormatting(after.Select(t => A1.Parse(t).ToNpoi()).ToArray(), rules);
                }
            }
            catch (Exception) { /* the ranges of an old-format sheet stay where they were rather than losing the edit itself */ }
        }
    }
}
