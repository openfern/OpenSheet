# Excel oracle tests

These tests ask a real Microsoft Excel what it thinks of OpenSheet's work. Excel defines the file format and the semantics of a spreadsheet,
so it is the right judge for everything that is not an API decision. They need **Windows with Excel installed**; without Excel every test marked
`[ExcelFact]` is skipped, and the two `*GoldenTests` replay Excel's recorded answers so the important checks still run on CI.

| File | What it checks |
|---|---|
| `ExcelOpensFilesTests` | files written by OpenSheet open without repair, and Excel sees every feature (and every chart) that was written; XLS; encryption; load → save |
| `FormulaOracleTests` (+ `FormulaBank`, `FormulaHarness`) | each formula of the bank gives Excel's result. A `Known` list holds the differences that are understood; a new one fails, a fixed one fails too |
| `FormulaGoldenTests` | the same, against `formula-bank.golden.txt`, without Excel |
| `NumberFormatOracleTests` | the display text of ~1,700 value/format pairs |
| `NumberFormatGoldenTests` | the same, against `numberformat.golden.txt`, without Excel |
| `OperationOracleTests` | insert/remove rows and columns, copy, sort — on XLSX and XLS — compared cell by cell |
| `StructureFeaturesOracleTests` | merged regions, links, comments, validations, conditional formats, names and tables follow those edits |
| `EditedFeaturesOracleTests` | features edited after creation (diagonal direction, validation operator, replaced link) and removed / deleted / replaced charts and sheets read the same in Excel, which opens the files |
| `FormulaBankCoverageTests` | every function the library adds is in the bank (needs no Excel) |
| `RoundTripOracleTests` | a workbook built by Excel keeps its features through OpenSheet's load → edit → save |

## Running

```
dotnet test tests/OpenSheet.Tests --filter FullyQualifiedName~Oracle
```

Reports with every disagreement are written to the temp folder: `opensheet-formula-oracle.txt`, `opensheet-numberformat-oracle.txt`,
`opensheet-operation-oracle.txt`, `opensheet-structure-oracle.txt`, `opensheet-roundtrip-oracle.txt` (and the two files of the round trip).

To record Excel's answers again (after adding formulas to the bank or formats to the matrix), on a machine with Excel:

```
$env:OPENSHEET_RECORD_GOLDEN = "1"
dotnet test tests/OpenSheet.Tests --filter "FullyQualifiedName~Evaluator_AgreesWithExcel|FullyQualifiedName~FormattedCellValue_MatchesWhatExcelDraws"
```

then commit the two `*.golden.txt` files.

## Things that bite

* **A dialog freezes the run.** Opening an encrypted file *without* a password makes Excel show a prompt that no setting suppresses. The tests never do that. If a run hangs, `Get-Process EXCEL,testhost | Stop-Process`.
* **Excel says "busy"** (`RPC_E_SERVERCALL_RETRYLATER`) while it starts, or when the default printer is a slow network queue. The oracle retries for up to four minutes.
* **The language of Excel leaks into results.** `Range.NumberFormat` is read in the user-interface language (a Spanish Excel does not take `yyyy` as the year), so formats are sent inside a saved file. `TEXT`, `DOLLAR`, `VERDADERO`, `Columna1` follow the language too; such cases are left out or normalized and say so.
* **`PageSetup` depends on the printer**, but only when the file asks for the printer's defaults, which OpenSheet never does.
* Excel keeps decimal and thousands separators in application options; the oracle sets `.` and `,` for the run and puts the user's values back when it quits.
