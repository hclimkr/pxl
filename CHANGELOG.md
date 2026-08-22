# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.9.5] - 2026-08-22

### Added

- **`java.util.UUID` is now a supported column type, in both directions and in all four paths** (Excel import,
  CSV import, Excel export, CSV export), including `Collection<UUID>`. Export was already possible through the
  custom-object path - a `UUID` has a `toString()` - but import failed while the column metadata was resolved,
  because `UUID` has no single-argument `String` constructor and, being a JDK type, cannot carry a
  `@PxlImportConverter` method either; there was no way for a caller to work around it short of a wrapper class.
  A `UUID` column is written as a string cell in the canonical lower-case form and read back from the canonical
  8-4-4-4-12 form in either case.
  Import is deliberately stricter than `UUID.fromString`, which counts the hyphen-separated groups but not their
  digits and so reads `"1-1-1-1-1"` as `00000001-0001-0001-0001-000000000001` - a typo silently becoming another
  identifier. The form is validated before the value is built, which also keeps the same file reading the same way
  on every JDK. The hyphen-less 32-digit form, a braced `{...}` form and a `urn:uuid:` prefix are refused as well,
  since export only ever writes the canonical form. `exportTrim` and `exportMasking` apply to a UUID column as they
  do to any other; `pattern`/`importPattern`/`exportPattern` has no meaning for a UUID and is ignored.
  For anyone already exporting a `UUID` through the custom-object path, the written text is unchanged.

### Changed

- **Some public `util/` methods now declare a checked exception they did not declare before, which is a
  source-breaking change** - the only one in this release. `PxlWorkbookUtils.createFormulaEvaluator`,
  `PxlSheetUtils.cloneSheet` and `PxlMiscUtils.convertColumnStringToColumnIndex` gain
  `throws PxlNullPointerException`; `PxlMiscUtils.convertCellReferenceStringToIndexes` and, through it,
  `PxlCellUtils.getCell(Sheet, String, boolean)` and the thirteen public helpers that address a cell by an A1-style
  reference gain it alongside the `PxlArgumentException` they already declared; and both
  `PxlCellUtils.addPicturesToCell` overloads gain `throws PxlArgumentException`. The index-argument sweep (see
  Fixed) adds `throws PxlArgumentException` to the rest of that surface: `PxlRowUtils.getRow`, `copyRow`,
  `copyRowMultiplyByRange`/`ByCount` and `removeRowsByRange`/`ByCount`; `PxlCellUtils.getCell(Sheet, int, int,
  boolean)`, the ten index-form `setCellValue` overloads, `setCellFormula`, `setCellErrorValue`, `setCellBlank` and
  the index form of `copyCell`; `PxlSheetUtils.setPrintArea(Sheet, int, int, int, int)`; and the three
  `PxlMiscUtils` index-to-text conversions. `getCellWithMerges` is deliberately not among them - it only ever
  reads, so its read path was split off to keep it free of the declaration. A call site that already catches
  `PxlException` compiles unchanged; one that catches only `PxlArgumentException`, or nothing at all, has to widen.
  Nothing moves for annotation-driven import/export - the core reaches none of these signatures from user code.

- **`@Valid` on a `@PxlSheet` field no longer gets the same row validated twice.** In the workbook form the binder
  validates each row on its own, tagging the violation with the sheet name and - on import - the row index. Marking
  the sheet field `@Valid` made bean validation walk those very rows a second time while validating the workbook
  object: the outcome was identical either way, so the extra traversal only cost time, and on export it cost
  diagnostics too, because the workbook pass runs first and carries no location - it reported untagged what the
  per-row pass would have pinned to a sheet. The validator is now configured to refuse that one cascade, leaving the
  per-row pass as the single place sheet rows are validated. Two consequences worth noting: a violation on export is
  now reported with its sheet name, and a sheet the binder does not process (`exportEnabled = false` on export,
  `importEnabled = false` on import) has its rows left unvalidated even when the field carries `@Valid`, so
  validation follows the sheets that are actually written or read - disabling a sheet through a runtime option
  behaves the same way. Constraints on the collection itself (`@NotEmpty`, `@Size`, ...), the workbook object's own
  constraints, and a `@Valid` on any field that is not a sheet are all unaffected, as is the sheet form, which never
  had a second pass to begin with.
- **A numeric or `java.util.Date` `pattern` now has to match the cell value in full.** A patterned column was
  parsed with `DecimalFormat.parse(String)` / `SimpleDateFormat.parse(String)`, which stop at the first character
  the pattern cannot read and report success with whatever came before it: under `"#,##0"`, `"123abc"` bound as
  `123` and `"1e3"` as `1`; under `"yyyy-MM-dd"`, `"2024-01-02 xxx"` bound as 2 January 2024. Naming a pattern
  therefore made a column *more* permissive than leaving it off, where `Integer.parseInt` and
  `LocalDate.parse(CharSequence, DateTimeFormatter)` have always required the whole string. Parsing now runs
  through a `ParsePosition` and rejects anything left over, so those values raise `PxlCellCodecException` instead
  of binding silently. Values a pattern reads end to end are unaffected, including prefixes and suffixes that are
  part of it (`"$1,234"` under `"$#,##0"`, `"50%"` under `"#0%"`). Two consequences worth noting: with
  `importTrim = false`, trailing whitespace (`"123 "`) is itself unconsumed input and is now rejected; and for
  `Date`, a custom pattern that matches only part of the value counts as a miss, so the built-in read formatters
  and the ISO-8601 instant still get their turn. Only consumption is checked - a misplaced grouping separator
  (`"1,2,3"`) still parses, as `DecimalFormat` does not verify group sizes.

### Fixed

- The index-taking `util/` helpers let a negative index through to POI, which answers one with a raw
  `IllegalArgumentException` - or, worse, with a value. `PxlRowUtils.getRow(sheet, -1, true)` and everything built
  on it (`copyRow`, `copyRowMultiplyByRange`/`ByCount`, `removeRowsByRange`/`ByCount`, and through
  `PxlCellUtils.getCell` the whole `setCellValue` family, `setCellFormula`, `setCellErrorValue`, `setCellBlank` and
  the index form of `copyCell`) reached `createRow(-1)`; `PxlSheetUtils.cloneSheet` and `setPrintArea` and
  `PxlCellUtils.addPicturesToCell` did the same for sheet, row and column indexes. The `PxlMiscUtils` conversions
  were the quiet ones: POI renders a negative column index as an empty string, and reads row or column `-1` as
  "not stated" rather than as an error, so `convertIndexesToCellReferenceString(-1, 0)` came back as `"A"` and
  `(0, -1)` as `"1"` - half a reference, which reads as a reference until it is used as one. This is the mirror of
  `convertColumnStringToColumnIndex` answering `"1"` with `-16`, fixed earlier in this release.
  The rule now is **lenient on reads, strict on writes**: a read (`createIfNone = false`, and `getCellWithMerges`,
  which only ever reads) answers a negative index with `null`, the way it answers an absent row or cell, while
  anything that changes the sheet - or builds a reference to hand back - raises `PxlArgumentException` naming the
  parameter. Only the lower bound is checked; an index past the format's last row or column is still POI's to
  refuse, since it names the limit in its message. Two consequences worth noting: those helpers now declare a
  checked `PxlArgumentException` (see Changed), and `PxlColumnUtils.autoSizeColumns` is left as it stands, its
  negative index still a documented no-op.
- `PxlCellUtils.cloneCellStyle` let a `null` argument out as a raw `NullPointerException`. Both overloads
  dereferenced what they were handed - `cloneCellStyle(Cell)` to reach the cell's own workbook,
  `cloneCellStyle(Cell, Workbook)` to create the style in the target - so a missing cell or target workbook left a
  public utility as a bare JDK exception instead of a `PxlException`. Both now answer `null`, which is already what
  the method returns for a style the target workbook has no room for, so a caller that handles "no style came back"
  needs no new branch. Returning rather than raising is deliberate: `copyCell` reaches this method from inside the
  `IntStream.forEach` lambda that `PxlRowUtils.copyRow` and `copyRowMultiplyByRange` copy cells with, where a
  checked `PxlNullPointerException` cannot be thrown, and both of those already treat a missing cell as a no-op.
  One unchecked path in `util/` is left as it stands: an out-of-range sheet index handed to
  `PxlSheetUtils.cloneSheet` still surfaces POI's `IllegalArgumentException`, since no index-taking helper checks
  its bounds and they are worth settling together rather than one at a time.
- `PxlCellUtils.cloneCellStyle` swallowed a style-creation failure without a word. When the target workbook has no
  room for another cell style POI raises `IllegalStateException`, and the helper caught it, discarded it and
  returned `null`; `copyCell` reads that `null` as "no style to set" and skips the call, so the cell was copied
  with its value and none of its formatting, and nothing anywhere said so. The reach is narrow - `copyCell` clones
  a style only across workbooks, since two cells in one workbook share the style object outright - but it is the
  quietest kind of loss inside it: rows past the limit come out unformatted and the file looks finished. The
  failure is now reported as an SLF4J `WARN` carrying POI's message, the way the export path already reports a
  style it cannot create, and the copy still goes ahead unstyled rather than aborting. A style created before the
  failure is discarded too, so a partial clone is no longer handed back as if it were a copy of the source.
- Public `util/` entry points let raw POI and JDK exceptions out instead of `PxlException`. Passing `null` to
  `PxlWorkbookUtils.createFormulaEvaluator`, `PxlSheetUtils.cloneSheet`,
  `PxlMiscUtils.convertColumnStringToColumnIndex` or `convertCellReferenceStringToIndexes` produced a bare
  `NullPointerException` from inside POI or the JDK, and `PxlCellUtils.addPicturesToCell` with
  `horizontalImageNum = 0` produced `ArithmeticException: / by zero` - it sizes the grid with
  `(imageNum + n - 1) / n` before it looks at whether there are any images at all, so even an empty list failed.
  Each is now checked with `PxlAssertSupport` and raises `PxlNullPointerException` or `PxlArgumentException`, the
  way the builders and the rest of `util/` already do. `convertColumnStringToColumnIndex` also stops trusting POI
  to reject what is not column letters: `CellReference.convertColStringToIndex` folds every character into the
  same running total and answers a nonsense index rather than failing (`"1"` came back as `-16`), so the shape is
  settled before the call and anything else raises `PxlArgumentException`.
- `PxlSheetUtils.cloneSheet` re-pointed only the first range of a print area at the clone. POI's two calls are not
  symmetric - `getPrintArea` names the sheet in front of every range, while `setPrintArea` wants the ranges bare and
  puts the destination sheet's name on each one itself - and the bridge between them cut the string at its first
  `!`. A print area of several ranges (`S!$A$1:$B$2,S!$D$1:$E$2`) therefore kept `S!` on every range after the
  first, leaving the clone pointing at the source sheet's cells wherever the resulting reference was accepted at
  all; and since `!` is not a character Excel forbids in a sheet name, a sheet named `A!B` has its name quoted
  rather than rejected, so the cut landed inside the quotes and broke the reference. Ranges are now separated on
  the commas outside a quoted name and each loses everything up to the first `!` outside the quotes, so a name
  holding a comma, a `!` or an escaped quote (`O'Brien`) comes through whole. The clone is also given its final
  name before the print area is set, because `setPrintArea` stamps the name the sheet carries at that moment. Only
  the public `util/` helper is affected - the import/export core does not call it.
- `PxlSheetUtils.cloneSheet` failed with POI's raw `IllegalArgumentException` when it was given a name the workbook
  already held. The requested name was only run through `WorkbookUtil.createSafeSheetName`, which replaces invalid
  characters and truncates to 31 chars but says nothing about uniqueness, and POI's `setSheetName` turns down a
  name another sheet holds, in whatever case it is asked in. The failure also came too late to undo anything: the
  clone had joined the workbook two calls earlier, so the caller was left with an extra sheet carrying POI's
  interim name and no print area, and with no `PxlException` to say so. The name is now made unique the way the
  names of the sheets an export creates are - a `" (2)"`, `" (3)"` ... suffix within the 31-char limit, compared
  ignoring case - so the clone succeeds and the name it ends up with is no longer guaranteed to be the one asked
  for; read it back with `workbook.getSheetName(workbook.getSheetIndex(clone))` where that matters. The name is
  settled before the clone joins the workbook, so the interim name POI gives the clone is not itself read as a
  collision, and a name no other sheet holds is applied unchanged. Only the public `util/` helper is affected - the
  import/export core does not call it.
- `PxlCellUtils.copyCell` moved a cell's comment and hyperlink instead of copying them. Both were passed to the
  destination's setters as the source cell's own objects, and POI re-anchors what it is handed, so the note and the
  link disappeared from the source cell. A second copy from the same source then found nothing left, which is what
  `PxlRowUtils.copyRow` / `copyRowMultiplyByRange` do per cell: replicating one template cell left the note on a
  single destination - the first for XLSX, the last for XLS, where an `HSSFCell` keeps answering with the reference
  it caches even though the sheet anchors only one note, so the loss showed up only in the saved file. Across
  workbooks the comment never reached the destination at all and instead moved to the destination's coordinates
  *inside the source workbook*, while the hyperlink ended up shared by both workbooks, so editing one changed the
  other. Both are now re-created as new objects owned by the destination sheet, carrying text, author, visibility,
  link type, address and label, with the comment's anchor keeping its source size over the destination cell; a
  comment's rich-text runs are flattened to plain text. Only the public `util/` helpers are affected - the
  import/export core does not call them.
- A `String` column with `exportStringAsPicture` wrote a value starting with `=` as quote-prefixed text instead of
  embedding a picture. The export checked the leading `=` before it looked at any option, so the picture branch was
  only reachable for values that did not start with one - and the same attribute on a `Collection` column had no
  such check, making one annotation behave differently by field type. The options now pick the form first, with
  `exportStringAsFormula` taking precedence over `exportStringAsPicture` when both are set; the quote prefix stays
  as what it always was, a safeguard on how a plain-text value is written.
- An empty CSV value bound a space (`0x20`) to a `char` column instead of leaving it at the Java default. A blank
  Excel cell never reaches the codec - the resolver drops it and the field keeps what it held - so the same absent
  value meant two different things depending on the source, and `char` was the only primitive where that was
  visible (`int` and `double` already yield `0`/`0.0`). An empty value now parses to `(char) 0` on both paths.
- A `boolean`/`Boolean` column read a numeric cell through `Math.abs(value) > 0.0000001` instead of comparing it
  with zero, so a genuine small value such as `1e-8` bound as `false` with nothing to show for it. The cutoff had
  no recorded rationale and contradicted the REFERENCE, which has always said a numeric cell is `true` when it is
  not 0. The comparison is now `value != 0`; both signed zeros stay `false`.
- `BigInteger` and `BigDecimal` columns with a `pattern` threw `ClassCastException` on the infinity and NaN
  tokens (`"∞"`, `"NaN"`). Their formatter runs with `setParseBigDecimal(true)`, which still returns a
  `Double` for those two, and the codec cast the result to `BigDecimal` unconditionally. The importer wrapped the
  failure with the cell coordinates, so nothing escaped, but the message named a cast rather than the value.
  Neither has a `BigDecimal` form, so both are now refused with the same diagnostic the `Double` and `Float`
  codecs use.
- Javadoc that contradicted the code: `PxlConstants.DEFAULT_EXPORT_IF_NULL` / `DEFAULT_EXPORT_IF_EMPTY` decide
  whether the **sheet** is created for a null or empty row collection, not whether a cell is written;
  `PxlImportSheetOption.importExcludeHiddenRows` / `importExcludeHiddenColumns` **exclude** hidden rows and columns
  rather than import them; and `PxlColumnSupport` named `exportDropdownList`, which does not exist, instead of
  `exportOptionItems` / `exportEnumDropDownListStyle`. Two `{@link}` references to `PxlSheet` in
  `PxlExportWorkbookOption` now resolve. Comments only — no signature, default or behavior changed.
- More of the same in `util/`, found by reading the package through after the index-argument sweep.
  `PxlCellUtils.getCellWithMerges` was documented as falling back to "the plain cell" on a streaming sheet; it
  cannot, since such a sheet reports neither merged regions nor a row looked up by index, so it yields `null`
  there - the class comment said the same thing and is corrected with it. Three descriptions were short of what
  the code does: the one-argument `getCellStringValue` did not say a blank cell also yields `null`, neither
  overload mentioned that an error cell does, and the fourteen A1-reference helpers documented
  `PxlArgumentException` only for a reference missing its row or column, when a blank one and one POI cannot read
  raise it too. The rest is wording: `cloneCellStyle` misnamed which call sits inside the lambda,
  `PxlSheetUtils`'s class comment opened on a "Both concerns" with no antecedent and described the streaming
  no-op too broadly, and `addNoteToCell` and `addPicturesToCell` did not mark their nullable `cell` parameter the
  way the rest of the class does. Comments only.

## [0.9.4] - 2026-08-11

### Added

- **CSV export.** `Pxl.exportCsv()` and `Pxl.exportSampleCsv()` write the same annotated objects the Excel export
  writes — same annotations, converters, column order, i18n names and bean validation. A CSV file holds one sheet,
  so both builders have the sheet form only (no `workbook(...)`) and terminate at `toFile(File)` /
  `toStream(OutputStream)`. The read axis (XLS · XLSX · CSV) and the write axis finally match.
- CSV export honors the coordinate attributes `exportHeaderRowIndex`, `exportFirstDataRowIndex`,
  `exportLastDataRowIndex`, `exportFirstDataColumnIndex` and `exportLastDataColumnIndex`, as CSV import already
  honors their `import*` counterparts. A row above the header, or a column before the first, is written as a record
  of **empty fields** rather than as a blank line, which PXL's own import would skip.
- CSV export ignores what CSV cannot carry: stylers, column widths, row heights, freeze panes, auto-filters,
  dropdowns, the engine, and `exportGroupingFieldName` (rows keep the order given). `exportStringAsFormula` writes
  the text verbatim, leading `=` and all, and `exportStringAsPicture` writes the image location; `exportPassword` is
  refused with `PxlArgumentException`, since CSV cannot be encrypted and writing plaintext would be a leak.
- A CSV export renders its whole output before the destination is opened, so a codec, validation or limit failure
  leaves no file behind and `toStream(...)` writes in one pass. The first
  `PxlConstants.EXPORT_MEMORY_THRESHOLD_OF_CSV` (4 MiB) is held in memory and the rest spills to a temporary file
  removed before the call returns, so the heap does not grow with the output — at the cost of disk space.
- `@PxlWorkbook` and `@PxlSheet` gain `exportCsvCharset`, `exportCsvDelimiter` and `exportCsvBom`, with matching
  fields on `PxlExportWorkbookOption` and `PxlExportSheetOption`, resolving through the same five levels as the
  import pair. A byte order mark is written only for UTF-8, UTF-16LE and UTF-16BE; any other charset drops it, as
  `UTF-16` has its encoder emit one already and EUC-KR and the like cannot encode U+FEFF at all.
- `PxlOptionalBoolean` (`UNSPECIFIED` · `TRUE` · `FALSE`), the type `exportCsvBom` takes on both annotations. An
  annotation element cannot hold `null`, so a `boolean` has no value left to mean "not specified" — its `false`
  would read as silence, and a sheet could never turn off a mark its workbook asked for. The option classes keep a
  boxed `Boolean`, whose `null` already carries that meaning.
- `@PxlSheet(importCsvCharset)` / `(importCsvDelimiter)` and the matching `PxlImportSheetOption` fields, so the
  sheets of one CSV workbook can be read with different encodings and delimiters — a workbook is one file per sheet,
  so a legacy MS949 export beside a UTF-8 one had no way to say so. Resolution runs sheet option → `@PxlSheet` →
  workbook option → `@PxlWorkbook` → built-in default; in the sheet form a wildcard option is the sheet-level route.

### Fixed

- A `null` row object in an exported collection is now a `PxlDataException` naming the sheet and the one-based
  position of the element (`core.export.rowNull`), raised before a row is created. It used to surface as a
  `PxlCellCodecException` blaming the **first column** — one that had done nothing, since no codec had run — with
  `java.lang.NullPointerException` as its entire message. The grouping branch no longer routes it to `(ungrouped)`.

### Changed

- `PxlFileFormat` and `PxlExcelEngine` moved from the base package to the new public package
  `io.github.hclimkr.pxl.type`, where `PxlOptionalBoolean` also lives — public enums the API takes as values, kept
  out of a root that is for the entry point and the shared defaults. **Update the import**; nothing else about
  either type changed: same constants, same lookups, same behavior.
- The CSV column cap is 16,384 rather than 100, on both `PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_COLUMNS` and
  `EXPORT_MAX_NUMBER_OF_CSV_COLUMNS`, so a row class written to XLSX can be read back from CSV. 100 had no basis in
  the format; nothing narrows, so no CSV that imported before stops importing. The row cap stays at 100,000. Neither
  cap guards memory: they bound counts, not bytes, and on import both are checked once the file is already loaded.
- **The six option classes are immutable and builder-only.** Every field of
  `Pxl{Import,Export}{Workbook,Sheet,Column}Option` is `final` and `@Setter`/`@NoArgsConstructor` are gone, which
  removes **65 setters and 6 no-argument constructors** that nothing called. `@Getter`, `@AllArgsConstructor`,
  `@Builder` and `add*Option(...)` stay; building a column option without `fieldName` now fails fast instead.
- `@PxlWorkbook(importCsvCharset)` / `(importCsvDelimiter)` default to the sentinels
  `PxlConstants.UNSPECIFIED_IMPORT_CSV_CHARSET` (`""`) and `UNSPECIFIED_IMPORT_CSV_DELIMITER` (`'\0'`) rather than
  to `"UTF-8"` and `','`, moving the effective defaults to the bottom of the cascade so a sheet can name them
  explicitly. The only observable change: a blank charset falls back to `"UTF-8"` instead of throwing.

## [0.9.3] - 2026-08-05

### Added

- `PxlExcelEngine` (`HSSF` / `XSSF` / `SXSSF`) names the POI implementation that writes a workbook, and each
  constant knows the `PxlFileFormat` it produces through `getFileFormat()`. It also carries the two lookups that
  belong on that axis: `fromWorkbookObject(Class)`, moved here from `PxlFileFormat`, and `fromPoiWorkbook(Workbook)`,
  which tells `XSSF` and `SXSSF` apart.

### Changed

- **Breaking.** `@PxlWorkbook(exportFileFormat)` is now `(exportExcelEngine)` and takes a `PxlExcelEngine`;
  `PxlExportWorkbookOption.exportFileFormat` is renamed the same way, and `PxlConstants` gains
  `DEFAULT_EXPORT_EXCEL_ENGINE` (`XSSF`) while `DEFAULT_EXPORT_FILE_FORMAT` becomes the physical `XLSX`. One enum
  had carried two questions — which writer runs, and what the bytes are — so an impossible writer no longer compiles:

  | Before                                       | After                                          |
  |----------------------------------------------|------------------------------------------------|
  | `@PxlWorkbook(exportFileFormat = ...)`       | `@PxlWorkbook(exportExcelEngine = ...)`        |
  | `PxlExportWorkbookOption.builder().exportFileFormat(...)` | `....exportExcelEngine(...)`       |
  | `PxlFileFormat.HSSF`                         | `PxlExcelEngine.HSSF`                          |
  | `PxlFileFormat.XSSF`                         | `PxlExcelEngine.XSSF`                          |
  | `PxlFileFormat.SXSSF`                        | `PxlExcelEngine.SXSSF`                         |
  | `PxlFileFormat.CSV` as an export target      | not expressible — CSV export is unsupported    |
  | `PxlConstants.DEFAULT_EXPORT_FILE_FORMAT` as the annotation default | `PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE` |
  | `PxlFileFormat.fromWorkbookObject(Class)`    | `PxlExcelEngine.fromWorkbookObject(Class)`     |

- **Breaking.** `PxlFileFormat` names physical formats only — `XLS`, `XLSX` and `CSV` in place of `HSSF`, `XSSF`,
  `SXSSF` and `CSV`. `XSSF` and `SXSSF` had held byte-identical extension, content type and limits all along, being
  one format written two ways. `fromPoiWorkbook(Workbook)` stays but answers on this axis, so a streaming-reader
  workbook now reports `XLSX`; ask `PxlExcelEngine.fromPoiWorkbook(...)` when the writer is what you need.
- Sheet names are now matched **ignoring case** on import, in both the Excel and the CSV reader: `Employees.csv`
  binds `@PxlSheet(name = "employees")`, and an Excel sheet named `EMPLOYEE` binds the alias `Employee`. A CSV sheet
  is named after its file, and a file name carries whatever casing the file system holds. Whitespace is still
  removed from both sides before comparing, and **column headers are unchanged: they remain case-sensitive**.
- The duplicate-sheet-name check on export ignores case too, so two sheets named `Employees` and `EMPLOYEES` are
  rejected with `PxlDataException` before anything is written, instead of reaching POI and surfacing as a
  `PxlSystemException` with sheets already created. It covers the sheet form, the sample export and the workbook
  form alike. No export that used to succeed fails now — a workbook could never hold both sheets.
- `importOverrideSuperClassSheet` / `exportOverrideSuperClassSheet` now recognize an override ignoring case as well,
  so a subclass field declared `@PxlSheet(name = "EMPLOYEES")` overrides a superclass sheet named `Employees`. The
  two names denote one sheet — on import both fields would otherwise bind the single sheet that matches either name,
  and on export a workbook cannot hold both.
- The contributing guide (`CONTRIBUTING.md` / `CONTRIBUTING_ko.md`) now states the repository's policy — issue
  reports and suggestions only, pull requests are not accepted — and asks for the version and artifact, expected
  versus actual behavior, and a minimal reproduction, with sensitive data stripped from any attached source file.
- `exportOptionItems` on a `String` column now goes through the workbook's content-i18n bundle, so the dropdown
  offers the same text the cells hold rather than the raw keys. A column of any other type writes its value in
  canonical form, so its items — and the enum constants used when no items are given — stay verbatim; a workbook
  without `exportI18nBaseName` is unaffected.
- Internal only: the shared base behind the export builders is now two layers. `PxlAbstractExportBuilder` keeps the
  export option and the `toFile`/`toStream` terminals, delegating the writing to three seams (`prepare()` before the
  destination is opened, `writeTo(OutputStream)`, `cleanup()` in the `finally`), while a POI-only
  `PxlAbstractExcelExportBuilder` holds `toWorkbook()` and the workbook side. No public signature moved.

### Fixed

- A numeric cell whose serial number is no Excel date — a negative one, for instance — is now rejected with a
  `PxlCellCodecException` naming the value. POI answers `null` for such a serial, which every java.time codec
  dereferenced into a message-less `NullPointerException`, while the `Date` codec bound it as no value at all;
  both paths now read the cell through one helper and report the same diagnostic.
- `exportSample` is now translated on `Collection<String>` and `Collection<Enum>` columns too. The i18n gate read
  the field type, so only String and enum scalars passed and a collection sample kept its raw bundle keys — or
  failed outright on `Collection<Enum>`, whose keys could not be parsed back into constants. Each element is now
  translated on its own, split by `exportCollectionSeparator`.

## [0.9.2] - 2026-07-29

### Added

- Excel import in the workbook form now names the workbook after its source file when `workbookName(...)` was not
  set: `importExcel().workbook(W.class).fromFile(file)` binds the file name without its extension (`Report.xlsx` →
  `"Report"`) to the `@PxlWorkbookName` field. An explicit name still wins, and `fromStream(...)` carries no file
  name, so it leaves the field `null` as before; a file that is nothing but an extension binds an empty name.
- `PxlFileFormat.fromPoiWorkbook(Workbook)`: resolves the file format of an already open POI workbook from its
  implementation type. `HSSFWorkbook` maps to `HSSF` and `SXSSFWorkbook` to `SXSSF`, while `XSSFWorkbook` and the
  streaming reader's `StreamingWorkbook` both map to `XSSF` — a streamed read opens the same OOXML container, and
  `SXSSF` denotes the streaming export workbook only.

### Changed

- **BREAKING** Every `sheet(...)` overload on the builders now leads with `rowClass`, and the parameters after it
  line up across builders — the collection second, the sheet name last: `exportExcel().sheet(Employee.class,
  employees, "Employees")`, `exportSampleExcel().sheet(Employee.class, "Employees")`,
  `importCsv().sheet(Employee.class, Set.class)`. Validation follows suit, so `rowClass` is reported first.
- **BREAKING** `Pxl.getWorkbookFileFormatFromWorkbookObject(Class)` moved to
  `PxlFileFormat.fromWorkbookObject(Class)`. The lookup itself is unchanged.
- **BREAKING** `Pxl.getWorkbookNameFromWorkbookObject(Object)` moved to
  `PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(Object)`. The lookup is unchanged, including the fail-safe
  `null` it returns for a missing or unreadable `@PxlWorkbookName` field.

## [0.9.1] - 2026-07-27

### Added

- Import builders: `override(...)` and `workbookName(...)` can now also be chained after `workbook(...)`/`sheet(...)`
  on the returned source step, matching the export builders where `override(...)` may appear anywhere before the
  final step. Chaining them before the parse-target configuration keeps working; the value set last wins.
- `PxlSystemException`: the exception a builder's final (execute) step wraps a failure PXL does not classify in — a
  checked I/O failure, an unexpected runtime failure from POI, and so on — keeping the original accessible through
  `getCause()`.

### Changed

- **BREAKING** `PxlException` is now `abstract` and can no longer be instantiated (`new PxlException(...)`).
  Catching it and subclassing it are unaffected, and `throws PxlException` remains a valid contract on the final
  (execute) steps. What surfaces there is always a concrete subtype: the matching one for a classified failure, and
  `PxlSystemException` for everything else.
- Failing to open an Excel source (a missing file, an unreadable or password-protected container, an unsupported
  format) now surfaces as `PxlIOException` rather than the bare base type, so callers can catch that case on its own.

## [0.9.0] - 2026-07-24

First public release.

### Added

- Import: XLSX, XLS, and CSV into Java objects.
- Export: Java objects into XLSX (default), XLS, and streaming XLSX (SXSSF), selectable via
  `@PxlWorkbook(exportFileFormat = ...)`.
- Around 30 built-in field-type codecs (numbers, `BigInteger`/`BigDecimal`, full `java.time` including
  zoned/offset/`Duration`/`Period`, enums, collections, and custom objects), with per-column custom converters.

[Unreleased]: https://github.com/hclimkr/pxl/compare/v0.9.5...HEAD
[0.9.5]: https://github.com/hclimkr/pxl/compare/v0.9.4...v0.9.5
[0.9.4]: https://github.com/hclimkr/pxl/compare/v0.9.3...v0.9.4
[0.9.3]: https://github.com/hclimkr/pxl/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/hclimkr/pxl/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/hclimkr/pxl/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/hclimkr/pxl/releases/tag/v0.9.0
