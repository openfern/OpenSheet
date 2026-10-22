using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.UDF;

namespace OpenSheet.Internal.Formulas
{
    /// <summary>An Excel error raised while computing a function; becomes the cell's error value.</summary>
    sealed class XlError : Exception
    {
        public XlError(ErrorEval error) { Error = error; }
        public ErrorEval Error { get; }
        public static XlError Value => new XlError(ErrorEval.VALUE_INVALID);
        public static XlError Num => new XlError(ErrorEval.NUM_ERROR);
        public static XlError Div0 => new XlError(ErrorEval.DIV_ZERO);
        public static XlError NA => new XlError(ErrorEval.NA);
    }

    delegate ValueEval Impl(Ctx c, ValueEval[] a);

    /// <summary>A rectangular block of evaluated values (one cell is a 1x1 grid).</summary>
    sealed class Grid
    {
        public Grid(ValueEval[,] cells) { Cells = cells; }
        public ValueEval[,] Cells { get; }
        public int Rows => Cells.GetLength(0);
        public int Columns => Cells.GetLength(1);
        public IEnumerable<ValueEval> Flat() { for (var r = 0; r < Rows; r++) for (var c = 0; c < Columns; c++) yield return Cells[r, c]; }
    }

    /// <summary>Argument reading with Excel's coercion rules. Errors are thrown as <see cref="XlError"/>.</summary>
    sealed class Ctx
    {
        public Ctx(int row, int col, int sheetIndex = -1, IEvaluationWorkbook workbook = null) { Row = row; Col = col; SheetIndex = sheetIndex; Workbook = workbook; }
        /// <summary>The workbook being evaluated (its sheets and cells), or null when the evaluator did not hand it over.</summary>
        public IEvaluationWorkbook Workbook { get; }
        public int Row { get; }
        public int Col { get; }
        /// <summary>Zero-based index of the sheet holding the formula, or -1 when the evaluator did not say.</summary>
        public int SheetIndex { get; }

        public ValueEval Scalar(ValueEval v)
        {
            try { return OperandResolver.GetSingleValue(v, Row, Col); }
            catch (EvaluationException e) { throw new XlError(e.GetErrorEval()); }
        }

        public static bool IsMissing(ValueEval v) => v is MissingArgEval;

        public double Num(ValueEval v)
        {
            v = Scalar(v);
            switch (v)
            {
                case NumberEval n: return n.NumberValue;
                case BoolEval b: return b.BooleanValue ? 1 : 0;
                case StringEval s: return ParseNumber(s.StringValue) ?? throw XlError.Value;
                case ErrorEval e: throw new XlError(e);
                default: return 0; // blank and omitted arguments
            }
        }

        public double Num(ValueEval[] a, int i, double ifMissing) => i < a.Length && !IsMissing(a[i]) ? Num(a[i]) : ifMissing;
        public int Int(ValueEval v) => (int)Math.Truncate(Num(v));
        public int Int(ValueEval[] a, int i, int ifMissing) => i < a.Length && !IsMissing(a[i]) ? Int(a[i]) : ifMissing;

        public string Str(ValueEval v)
        {
            v = Scalar(v);
            switch (v)
            {
                case StringEval s: return s.StringValue;
                case NumberEval n: return NumberText(n.NumberValue);
                case BoolEval b: return b.BooleanValue ? "TRUE" : "FALSE";
                case ErrorEval e: throw new XlError(e);
                default: return "";
            }
        }

        public bool Bool(ValueEval v)
        {
            v = Scalar(v);
            switch (v)
            {
                case BoolEval b: return b.BooleanValue;
                case NumberEval n: return n.NumberValue != 0;
                case StringEval _: throw XlError.Value; // Excel does not read "TRUE" or "1" as a yes/no argument
                case ErrorEval e: throw new XlError(e);
                default: return false;
            }
        }

        /// <summary>The cells of a range or array argument; a plain value becomes a 1x1 grid.</summary>
        public Grid Area(ValueEval v)
        {
            if (v is AreaEval area)
            {
                var rows = area.LastRow - area.FirstRow + 1; var cols = area.LastColumn - area.FirstColumn + 1;
                var cells = new ValueEval[rows, cols];
                for (var r = 0; r < rows; r++) for (var c = 0; c < cols; c++) cells[r, c] = Unwrap(area.GetRelativeValue(r, c));
                return new Grid(cells);
            }
            return new Grid(new[,] { { Unwrap(Scalar(v)) } });
        }

        static ValueEval Unwrap(ValueEval v) => v is RefEval r ? r.GetInnerValueEval(r.FirstSheetIndex) : v;

        /// <summary>
        /// Numbers for a statistical function: inside ranges only real numbers count (text, booleans and blanks are skipped),
        /// while a number given directly as an argument may be text or a boolean.
        /// </summary>
        public List<double> Numbers(IEnumerable<ValueEval> args)
        {
            var list = new List<double>();
            foreach (var a in args)
            {
                if (a is AreaEval)
                {
                    foreach (var x in Area(a).Flat())
                    {
                        if (x is NumberEval n) list.Add(n.NumberValue);
                        else if (x is ErrorEval e) throw new XlError(e);
                    }
                }
                else if (a is RefEval)
                {
                    var x = Unwrap(Scalar(a));
                    if (x is NumberEval n) list.Add(n.NumberValue);
                    else if (x is ErrorEval e) throw new XlError(e);
                }
                else if (!IsMissing(a)) list.Add(Num(a));
            }
            return list;
        }

        public List<double> Numbers(ValueEval[] a, int from = 0) => Numbers(a.Skip(from));

        /// <summary>Like <see cref="Numbers(IEnumerable{ValueEval})"/> but keeps text as 0 and booleans as 1/0 inside ranges (the "A" functions).</summary>
        public List<double> NumbersWithText(IEnumerable<ValueEval> args)
        {
            var list = new List<double>();
            foreach (var a in args)
            {
                // typed straight into the formula, text has to read as a number; inside a range or from a cell it counts as 0
                if (a is StringEval typed) { list.Add(ParseNumber(typed.StringValue) ?? throw XlError.Value); continue; }
                foreach (var x in a is AreaEval ? Area(a).Flat() : new[] { Unwrap(Scalar(a)) })
                {
                    if (x is NumberEval n) list.Add(n.NumberValue);
                    else if (x is BoolEval b) list.Add(b.BooleanValue ? 1 : 0);
                    else if (x is StringEval) list.Add(0);
                    else if (x is ErrorEval e) throw new XlError(e);
                }
            }
            return list;
        }

        public static double? ParseNumber(string s)
        {
            if (s == null) return null;
            s = s.Trim();
            if (s.Length == 0) return null;
            var percent = s.EndsWith("%", StringComparison.Ordinal);
            if (percent) s = s.Substring(0, s.Length - 1).TrimEnd();
            if (s.StartsWith("$", StringComparison.Ordinal)) s = s.Substring(1);
            if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d) && !double.IsNaN(d) && !double.IsInfinity(d))
                return percent ? d / 100 : d;
            return null;
        }

        public static string NumberText(double d) => d.ToString("G15", CultureInfo.InvariantCulture);
    }

    /// <summary>One implementation usable both as a built-in slot function and as a user-defined (add-in style) function.</summary>
    sealed class PackFunction : FreeRefFunction, NPOI.SS.Formula.Functions.Function
    {
        readonly Impl _impl;
        public PackFunction(Impl impl) { _impl = impl; }
        public Impl Impl => _impl;

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) => Run(args, ec.RowIndex, ec.ColumnIndex, ec.SheetIndex, ec.GetWorkbook());
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) => Run(args, srcRowIndex, srcColumnIndex, -1, null);

        ValueEval Run(ValueEval[] args, int row, int col, int sheet, IEvaluationWorkbook workbook)
        {
            try { return _impl(new Ctx(row, col, sheet, workbook), args); }
            catch (XlError e) { return e.Error; }
            catch (EvaluationException e) { return e.GetErrorEval(); }
            catch (OverflowException) { return ErrorEval.NUM_ERROR; }
            catch (DivideByZeroException) { return ErrorEval.DIV_ZERO; }
            catch (ArgumentException) { return ErrorEval.VALUE_INVALID; }
            catch (FormatException) { return ErrorEval.VALUE_INVALID; }
        }
    }

    /// <summary>Finds functions that NPOI 2.7.6 does not ship. Plugged into every workbook as a tool pack.</summary>
    sealed class PackFinder : UDFFinder
    {
        public static readonly PackFinder Instance = new PackFinder();

        public override FreeRefFunction FindFunction(string name)
        {
            if (name == null) return null;
            if (name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)) name = name.Substring(6);
            return FunctionPack.TryGet(name, out var f) ? f : null;
        }
    }

    static partial class FunctionPack
    {
        static readonly Dictionary<string, PackFunction> Map = new Dictionary<string, PackFunction>(StringComparer.OrdinalIgnoreCase);
        static readonly object Gate = new object();
        static volatile bool _ready;

        static void Add(string name, Impl impl) => Map[name] = new PackFunction(impl);

        public static bool TryGet(string name, out FreeRefFunction f)
        {
            EnsureReady();
            if (Map.TryGetValue(name, out var p)) { f = p; return true; }
            f = null; return false;
        }

        public static IReadOnlyCollection<string> Names { get { EnsureReady(); return Map.Keys.ToList(); } }

        /// <summary>Registers everything with NPOI. Safe to call from any thread, any number of times.</summary>
        public static void Initialize() => EnsureReady();

        static void EnsureReady()
        {
            if (_ready) return;
            lock (Gate)
            {
                if (_ready) return;
                Define();
                // Everything below reaches into NPOI's static tables. If a different NPOI layout makes that fail, the library still works,
                // only without the extra functions and corrections.
                try { RegisterBuiltinSlots(); InstallIntoNpoi(); } catch (Exception) { /* keep going with NPOI's own functions */ }
                _ready = true;
            }
        }

        // Names NPOI knows as built-ins but has no code for: fill the empty slot (the API refuses to replace working ones).
        static void RegisterBuiltinSlots()
        {
            foreach (var kv in Map)
            {
                try { FunctionEval.RegisterFunction(kv.Key, kv.Value); }
                catch (Exception) { /* not a built-in name, or already implemented by NPOI: the tool pack / NPOI serves it */ }
            }
        }

        /// <summary>Call before evaluating so the names above are known to the evaluator.</summary>
        public static void Attach(NPOI.SS.UserModel.IWorkbook book)
        {
            EnsureReady();
            book.AddToolPack(PackFinder.Instance);
        }
    }
}
