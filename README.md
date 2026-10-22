# OpenSheet

A free spreadsheet library for .NET: read, write and edit **XLSX, XLS, CSV/TSV, JSON and XML**, with formulas, styles,
data validation, conditional formatting, images, charts, tables and file encryption.

* **License:** Apache-2.0. No license key, no telemetry.
* **Targets:** `netstandard2.0` (.NET Framework 4.7.2+, older .NET), `net8.0` and `net10.0`. Windows, Linux, macOS.
* **Engine:** [NPOI](https://github.com/nissl-lab/npoi) 2.7.6, pinned on purpose (see [Licensing](#licensing-and-provenance)).
* **API:** `WorkBook` → `WorkSheet` → `Range` → `Cell`, all indexes zero-based. The shape follows an established commercial library so
  existing code ports with little change; it is an independent implementation (see `CLEANROOM.md`).

> Status: complete first version. 441 tests pass, the output validates against the OOXML schema, **99.7% of the reference API surface is present**, and the
> behavior that Excel defines (formulas, display text, edits, files Excel must open) is checked against a real Excel, not just against NPOI.
> It has **not** been compared with the reference's runtime behavior (it was never run, by design). Read [Limits](#limits).

## Quick start

```csharp
using OpenSheet;
using OpenSheet.Styles;

var book = WorkBook.Create();                       // XLSX; WorkBook.Create(ExcelFileFormat.XLS) for XLS
var sheet = book.DefaultWorkSheet;

sheet["A1"].Value = "Item";      sheet["B1"].Value = "Price";
sheet["A2"].Value = "Pen";       sheet["B2"].Value = 1.5;
sheet["A3"].Value = "Notebook";  sheet["B3"].Value = 4.25;
sheet["B4"].Formula = "=SUM(B2:B3)";

sheet["A1:B1"].Style.Font.Bold = true;
sheet["B2:B4"].FormatString = "#,##0.00";
sheet["A1:B1"].Style.SetBackgroundColor("#FFE699");

book.SaveAs("prices.xlsx");                         // format comes from the extension
book.SaveAs("prices.csv");                          // .xls .tsv .json .xml .html also work
book.SaveAs("secret.xlsx", "password");             // AES-256 encryption (XLSX only)

var again = WorkBook.Load("prices.xlsx");           // detects XLSX / XLS / encrypted / CSV / JSON / XML by content
decimal total = again.DefaultWorkSheet["B2:B3"].Sum();   // 5.75

// very large .xlsx files: forward-only, one row in memory at a time
foreach (var row in WorkBook.StreamRows("prices.xlsx", new StreamReadOptions { HasHeaderRow = true }))
    System.Console.WriteLine($"{row["Item"].Value}: {row["Price"].Value}");
```

More: ranges (`sheet["A1:C10"]`, `Rows`, `Columns`, `Sort*`, `Insert*`, `Remove*`), `DataValidations`, `ConditionalFormatting`,
`InsertImage`, `CreateChart`, `AddNamedRange`, `SaveAsNamedTable`, `Metadata`, `ToDataTable`/`ToDataSet`.

## What is in it

| Area | Notes |
|---|---|
| Formats | XLSX, XLS, CSV, TSV, JSON, XML in; the same plus HTML out. Format is detected from content when loading and chosen from the extension when saving. |
| Cells & ranges | typed values, aggregates (`Sum/Min/Max/Avg`), copy, replace, sort, insert/remove rows and columns, merge, freeze panes, grouping, breaks, print setup, headers/footers. References in formulas follow the cells exactly as in Excel. Whole-column and whole-row addresses (`"A:A"`, `"1:1"`) work |
| Formulas | NPOI's engine plus ~165 functions it lacks and 55 it answers differently from Excel at the edges; unsupported ones read as `#NAME?` (never 0); results are stored when saving so other readers see them. See [`spec/FORMULAS.md`](spec/FORMULAS.md) |
| Display text | `FormattedCellValue` draws General, dates, times and number formats like Excel (invariant culture) |
| Rules | data validation (list, integer, decimal, date, text length), conditional formatting (cell value, formula); both follow row and column edits |
| Drawing | pictures (JPEG, PNG, GIF, BMP, TIFF) and charts (column, bar, line, area, pie, scatter, XLSX only) |
| Tables & names | named ranges; named tables (XLSX), which grow, shrink and move with the cells; names are checked the way Excel checks them |
| Security | AES-256 workbook password; remembered when loading so saving does not silently remove it |

## Performance

Text is built with [ZString](https://github.com/Cysharp/ZString) (Cysharp, MIT): CSV/TSV is encoded to UTF-8 as it is written, and HTML, `Range.StringValue`
and date formats fill pooled buffers instead of chaining strings. NPOI already depends on ZString, so your dependency graph does not grow. Most of the rest
is work that was done twice:

* `sheet["A1"]` no longer walks every row of the sheet to work out the used range, which made writing *N* cells by address O(*N*²). A1 addresses are parsed
  by a scanner instead of a regex, and column names are cached.
* Streaming reads use `XmlReader` directly, with no node tree per row.
* A cell's number format is looked up once per style instead of once per cell; date format codes are parsed once.
* Saving evaluates formulas in one pass; exports fetch each row once and no longer go through a string on the way to bytes.

BenchmarkDotNet, ShortRun on .NET 10, a 20,000 × 8 sheet (5,000 rows for the write and formula cases), one run each, so read the ratios, not the digits:

| Workload | Before | After |
|---|---:|---:|
| Write 15,000 cells by address (`ws["A" + r].Value = …`) | 16,640 ms · 15.5 GB | 27 ms · 16 MB |
| Export CSV | 307 ms · 130 MB | 125 ms · 49 MB |
| Export JSON | 209 ms · 68 MB | 51 ms · 35 MB |
| Export XML | 377 ms · 120 MB | 122 ms · 61 MB |
| Export HTML | 434 ms · 134 MB | 87 ms · 59 MB |
| Stream 20,000 rows (`StreamRows`) | 248 ms · 129 MB | 80 ms · 37 MB |
| Stream with a header row | 358 ms · 129 MB | 79 ms · 37 MB |
| Load CSV | 1,429 ms · 217 MB | 296 ms · 203 MB |
| `Range.Sum()` over 160,000 cells | 270 ms · 70 MB | 80 ms · 41 MB |
| `ToDataTable` | 278 ms · 50 MB | 62 ms · 27 MB |
| `Range.StringValue` (16,000 cells) | 36 ms · 21 MB | 8 ms · 4 MB |
| `FormatString` on 48,000 cells | 97 ms · 21 MB | 57 ms · 1.5 MB |
| Save 5,000 formulas | 346 ms · 60 MB | 124 ms · 56 MB |

Saving, loading and the cell objects themselves are NPOI's, and cost what they cost. Reproduce with
`dotnet run -c Release -f net10.0 --project benchmarks/OpenSheet.Benchmarks -- --inprocess --job short --filter "*"`.

Things that behave differently from before, all of them corrections:

* The XML export declares `encoding="utf-8"`. It used to say `utf-16` over UTF-8 bytes, which XML parsers reject.
* Importing XML that has a `<!DOCTYPE>` is refused (external entities and entity-expansion bombs); the files OpenSheet writes have none.
* A date-formatted number that no date can hold (negative, or past year 9999) reads as a number instead of throwing, in memory and when streaming.
* Streaming rejects a shared-string index the package does not have with an `InvalidDataException`, and caps the size of the small parts it loads whole.

## Limits

* **Charts:** XLSX only; the series of a chart loaded from a file cannot be edited (title, legend and position can).
* **XLS:** old format limits apply; colors snap to the 56-color palette; password protection is not available; validations, conditional formats, links, comments and names follow row/column edits (checked against Excel for the cases in the oracle tests) basis.
* **Not supported:** creating pivot tables, macros (a `.xlsm` is written without them), form controls, slicers, sparklines. Pivot tables, sparklines, shapes and pictures that already exist in a loaded `.xlsx` are kept.
* **Formulas:** no dynamic arrays (`FILTER`, `SORT`, …) and no array arithmetic inside functions (`SUMPRODUCT((A1:A9>2)*…)`); those read as `#NAME?`/`#VALUE!` here and Excel computes them on open. `=-2^2` follows NPOI's precedence. Full list: [`spec/FORMULAS.md`](spec/FORMULAS.md).
* **Display text:** engineering notation, fraction padding and the accounting `*` fill differ from Excel.
* **Threads:** a `WorkBook` is not thread-safe (separate workbooks on separate threads are fine).
* **Other readers:** files open in Excel without repair (tested); LibreOffice and other consumers were not tested.

## Build, test, check

```
dotnet build OpenSheet.slnx                # library (netstandard2.0, net8.0, net10.0), tests and benchmarks; needs the .NET 10 SDK
dotnet test tests/OpenSheet.Tests          # schema validation with Microsoft's OOXML SDK, plus Excel itself when it is installed
.\tools\parity.ps1                         # compares the public API with spec/api-surface.md (PowerShell)
dotnet pack src/OpenSheet -c Release
```

The tests run on net8.0 and net10.0 everywhere, and on Windows also on .NET Framework 4.8, which exercises the `netstandard2.0` build that older runtimes get.
Versions are pinned in `Directory.Packages.props` and the SDK in `global.json`.

The tests under `tests/OpenSheet.Tests/Oracle` drive a hidden Excel through COM and are **skipped when Excel is not installed**. Excel's answers for the
formula bank and the number formats are recorded in `Oracle/*.golden.txt`, and replayed by tests that need no Excel. To record them again on a
machine with Excel: set `OPENSHEET_RECORD_GOLDEN=1` and run the two oracle tests. Details in [`tests/OpenSheet.Tests/Oracle/README.md`](tests/OpenSheet.Tests/Oracle/README.md).

A repository-local `nuget.config` points at nuget.org so a machine-wide package-source mapping cannot block restoring NPOI.

## Licensing and provenance

OpenSheet is Apache-2.0 (`LICENSE`, `NOTICE`).

* **NPOI is pinned to `[2.7.6]`.** NPOI 2.8.0+ ships its NuGet binaries under an "Open Source Maintenance Fee" EULA (revenue ≥ US$10,000/year pays a monthly fee). A library that depended on it would pass that obligation to every user. The NPOI *source* remains Apache-2.0; building NPOI yourself is the way to use 2.8 without the EULA.
* NPOI 2.7.6 depends on `NSax` (LGPL-3.0-only); see `NOTICE`.
* `spec/STUDY.md` records how the reference was studied; `CLEANROOM.md` the rules and honest limits of the process; `spec/DECISIONS.md` every place we had to choose.
* The reference product's EULA contains a clause about not competing with its functionality. **If you accepted that EULA, get legal advice before publishing this project.**
