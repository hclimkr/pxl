# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `@PxlSheet(importCsvCharset)` / `(importCsvDelimiter)`, and the matching `importCsvCharset` /
  `importCsvDelimiter` fields on `PxlImportSheetOption`, so the sheets of one CSV workbook can be read with
  different encodings and delimiters. A CSV workbook is read as one file per sheet, which put both attributes on
  the wrong level: a workbook holding a legacy MS949 export alongside a UTF-8 one had no way to say so. Resolution
  runs sheet option → `@PxlSheet` → workbook option → `@PxlWorkbook` → built-in default, and a sheet that names
  neither inherits the workbook value. Both are ignored for an Excel source, as the workbook-level pair already
  was. The sheet form (`sheet(...)`) binds no `@PxlSheet` field, so there a wildcard `PxlImportSheetOption` is the
  sheet-level route.

### Fixed

- A CSV import configured with an unusable charset or delimiter now fails with `PxlArgumentException` naming
  the attribute, where it used to surface as `PxlSystemException` naming neither. `Charset.forName(...)` and
  the delimiter check inside `CSVFormat.Builder.build()` both reject their input with unchecked exceptions,
  and both sat inside a `try` that catches `IOException` only — so `importCsvCharset("UTF8-typo")` or
  `importCsvDelimiter('\n')` (or `'"'`, which collides with the quote character) escaped to the builder
  boundary and were flattened by its catch-all. The two calls are now made before that block and normalized
  individually, keeping the original exception as the cause.

### Changed

- `PxlImportWorkbookOption.importResourceBundle` is now `final` like every other field of that class, so Lombok
  no longer generates `setImportResourceBundle(ResourceBundle)` for it. The field was the one non-final member
  left, and the setter it produced was the only one the class had. Build the option through its builder
  (`PxlImportWorkbookOption.builder().importResourceBundle(bundle)`), which is how the option is documented and
  how every other field was already set. `PxlExportWorkbookOption` keeps its setters for now.
- `@PxlWorkbook(importCsvCharset)` / `(importCsvDelimiter)` now default to the "not specified" sentinels
  `PxlConstants.UNSPECIFIED_IMPORT_CSV_CHARSET` (`""`) and `UNSPECIFIED_IMPORT_CSV_DELIMITER` (`'\0'`) rather than
  to `"UTF-8"` and `','`; the effective defaults moved to the bottom of the cascade above. Every value a caller
  would actually write behaves as before — the one difference is that an explicitly blank charset now falls back
  to `"UTF-8"` where it used to reach `Charset.forName("")` and throw. The sentinel is what lets a sheet name
  `"UTF-8"` or `','` explicitly to return to the default against a workbook that names something else; had the
  annotation default stayed a usable value, a sheet saying nothing and a sheet saying `"UTF-8"` would have been
  indistinguishable, and the workbook attribute would have been unreachable in the workbook form.
- The two CSV configuration errors now name the sheet they were resolved for:
  `core.import.csv.charsetInvalid` gained a leading `{0}`=sheetName, shifting the charset name to `{1}`, and
  `core.import.csv.delimiterInvalid` gained `{0}`=sheetName. With the values resolved per sheet, a workbook-wide
  message would leave the caller to guess which of the files is misconfigured.
- `PxlCellResolver.buildDataCell` returns the string it wrote rather than `void`, and computes that string
  without writing when given a `null` cell; `buildDataString(value, columnMeta)` exposes the cell-less call.
  The Excel path is untouched — its two call sites ignore the return value, and all 30 codecs already
  returned their string and guarded the cell. `internal/codec` is not public API.
- Diagnostic message keys are now grouped by binding direction. A key thrown on only one side starts with
  that direction — `builder.export.*`, `core.import.*`, `meta.export.*`, `codec.import.*` — and the format,
  where there is one, follows it (`core.import.csv.fileNameCountMismatch`). Keys genuinely shared by both
  sides keep their neutral name: `core.sheet.countExceeded` is thrown by the exporter and both importers,
  and `codec.columnType.unsupported` by all three codec entry points. These keys live in `internal/i18n`
  and are not public API, so calling code is unaffected; what the grouping buys is that a message can no
  longer sit under a name that hides where it comes from, which had already happened —
  `builder.workbookSheetExclusive` and `meta.maskingInvalid` both read as neutral although neither is
  reachable while importing.
- Codec parse failures now name the direction they came from, and the export wording changed with them.
  Exporting parses strings too: a `String` field value or an `exportSample` is parsed into the target type
  before being written back out, so `codec.parse.invalid` and the enum/object parse keys were thrown from
  both sides under a single name. Each is split in two (`codec.import.parse.invalid` /
  `codec.export.parse.invalid`, and likewise for `enum.parseFailed`, `enum.parseError`,
  `object.parseFailed`, `object.parseError`). **The export messages now open with "the export value '…'"**
  (Korean "출력할 값 '…'") so a failure while writing is no longer worded as one while reading. Import
  messages are unchanged.

## [0.9.3] - 2026-08-05

### Added

- `PxlExcelEngine` (`HSSF` / `XSSF` / `SXSSF`) names the POI implementation that writes a workbook, and
  each constant knows the `PxlFileFormat` it produces through `getFileFormat()`. It also carries the two
  lookups that belong on that axis: `fromWorkbookObject(Class)`, moved here from `PxlFileFormat`, and
  `fromPoiWorkbook(Workbook)`, which tells `XSSF` and `SXSSF` apart.

### Changed

- **Breaking.** `@PxlWorkbook(exportFileFormat)` is now `@PxlWorkbook(exportExcelEngine)` and takes a
  `PxlExcelEngine` instead of a `PxlFileFormat`; `PxlExportWorkbookOption.exportFileFormat` is renamed the
  same way, and `PxlConstants.DEFAULT_EXPORT_FILE_FORMAT` (`XSSF`) is joined by `DEFAULT_EXPORT_EXCEL_ENGINE`
  (`XSSF`) while itself becoming the physical default (`XLSX`). One enum used to carry two unrelated
  questions — which writer runs, and what the bytes are — which is what let a workbook declare
  `exportFileFormat = CSV`, a format no writer can produce, and fail at runtime with an unsupported-format
  `PxlDataException`. Naming a writer that cannot exist is now a compile error, so the check, its message and
  the `throws` on workbook creation are gone rather than reworded. Migration is mechanical:

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

- **Breaking.** `PxlFileFormat` now names physical formats only — `XLS`, `XLSX` and `CSV` in place of `HSSF`,
  `XSSF`, `SXSSF` and `CSV`. `XSSF` and `SXSSF` had held byte-identical extension, content type and limits all
  along, because they are one format written two ways; they are now the single `XLSX`. What the enum keeps is
  what a caller reads rather than declares: the filename extension and MIME content type for a download
  response, and the sheet/row/column limits an export is bound by. `fromPoiWorkbook(Workbook)` stays but
  answers on this axis, so a streaming-reader workbook now reports `XLSX` rather than `XSSF`; ask
  `PxlExcelEngine.fromPoiWorkbook(...)` when the writer is what you need.

- Sheet names are now matched **ignoring case** on import, in both the Excel and the CSV
  reader: `Employees.csv` binds a sheet declared as `@PxlSheet(name = "employees")`, and an
  Excel sheet named `EMPLOYEE` binds the alias `Employee`. A sheet name is not always typed
  by hand where the binding is declared — a CSV sheet is named after its file, and a file
  name carries whatever casing the file system holds, which on Windows is not distinguished
  at all. Whitespace is still removed from both sides before comparing, and **column headers
  are unchanged: they remain case-sensitive**.

- The duplicate-sheet-name check on export ignores case too, so two sheets named `Employees` and
  `EMPLOYEES` are rejected with `PxlDataException` before anything is written, instead of
  reaching POI and surfacing as a `PxlSystemException` with sheets already created. It covers
  the sheet form, the sample export, and the workbook form alike, and the message now carries
  the offending names. No export that used to succeed fails now — a workbook could never hold
  both sheets.

- `importOverrideSuperClassSheet` / `exportOverrideSuperClassSheet` now recognize an override
  ignoring case as well, so a subclass field declared as `@PxlSheet(name = "EMPLOYEES")`
  overrides a superclass sheet named `Employees`. The two names denote one sheet — on import
  both fields would otherwise bind the single sheet that matches either name, and on export a
  workbook cannot hold both.

- The contributing guide (`CONTRIBUTING.md` / `CONTRIBUTING_ko.md`) now states the
  repository's policy — issue reports and suggestions only, pull requests are not
  accepted — and asks for the version and artifact, expected versus actual behavior, and
  a minimal reproduction, with sensitive data stripped from any attached source file.

- `exportOptionItems` on a `String` column now goes through the workbook's content-i18n bundle, so the
  dropdown offers the same text the cells hold rather than the raw keys. A column of any other type writes
  its value in canonical form, so its items — and the enum constants used when no items are given — stay
  verbatim; a workbook without `exportI18nBaseName` is unaffected.

- Internal only: the shared base behind the export builders is now two layers. `PxlAbstractExportBuilder` keeps
  the export option and the `toFile(File)` / `toStream(OutputStream)` terminals — resource handling and exception
  normalization — and delegates the writing itself to three seams (`prepare()` before the destination is opened,
  `writeTo(OutputStream)`, `cleanup()` in the terminal's `finally`), while a new POI-only
  `PxlAbstractExcelExportBuilder` holds `toWorkbook()`, the workbook creation result, and the workbook
  implementation of those seams. Both classes are package-private and no public signature moved, so nothing
  changes for callers; the split is what lets a non-Excel writer reuse the terminals. The seam order is the
  contract: preparing before the destination is opened keeps a failed export from leaving an empty file, and
  releasing in the `finally` keeps a workbook (and, with `SXSSF`, its temp files) from being left behind when
  opening the destination fails.

### Fixed

- A numeric cell whose serial number is no Excel date — a negative one, for instance — is now rejected with a
  `PxlCellCodecException` naming the value. POI answers `null` for such a serial, which every java.time codec
  dereferenced into a message-less `NullPointerException`, while the `Date` codec bound it as no value at all;
  both paths now read the cell through one helper and report the same diagnostic.

- `exportSample` is now translated on `Collection<String>` and `Collection<Enum>` columns too. The i18n gate
  read the field type, so only String and enum scalars passed and a collection sample kept its raw bundle
  keys — or failed outright on `Collection<Enum>`, whose keys could not be parsed back into constants. Each
  element is now translated on its own, split by `exportCollectionSeparator`.

## [0.9.2] - 2026-07-29

### Added

- Excel import in the workbook form now names the workbook after its source file when
  `workbookName(...)` was not set: `importExcel().workbook(W.class).fromFile(file)` binds
  the file name without its extension (`Report.xlsx` → `"Report"`) to the
  `@PxlWorkbookName` field. An explicitly configured name still wins, and a stream source
  carries no file name, so `fromStream(...)` leaves the field `null` as before. A file that
  is nothing but an extension (`".xlsx"`) binds an empty name.

- `PxlFileFormat.fromPoiWorkbook(Workbook)`: resolves the file format of an already open
  POI workbook from its implementation type. `HSSFWorkbook` maps to `HSSF` and
  `SXSSFWorkbook` to `SXSSF`, while `XSSFWorkbook` and the streaming reader's
  `StreamingWorkbook` both map to `XSSF` — a streamed read opens the same OOXML
  container, and `SXSSF` denotes the streaming export workbook only.

### Changed

- **BREAKING** Every `sheet(...)` overload on the builders now leads with `rowClass`, and
  the parameters after it line up across builders — the collection second, the sheet name
  last — so a given position always carries the same meaning:
  `exportExcel().sheet(Employee.class, employees, "Employees")`,
  `exportSampleExcel().sheet(Employee.class, "Employees")`,
  `importExcel().sheet(Employee.class, Set.class, "Employees")` and
  `importCsv().sheet(Employee.class, Set.class)`. The import overloads that already led
  with `rowClass` are unchanged: `importExcel().sheet(Class, String...)`,
  `importExcel().sheet(Class, List)` and `importCsv().sheet(Class)`. Argument validation
  follows the new order, so `rowClass` is now the first argument reported as invalid.
- **BREAKING** `Pxl.getWorkbookFileFormatFromWorkbookObject(Class)` moved to
  `PxlFileFormat.fromWorkbookObject(Class)`. The lookup itself is unchanged.
- **BREAKING** `Pxl.getWorkbookNameFromWorkbookObject(Object)` moved to
  `PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(Object)`. The lookup itself is
  unchanged, including the fail-safe `null` it returns for a missing or unreadable
  `@PxlWorkbookName` field.

## [0.9.1] - 2026-07-27

### Added

- Import builders: `override(...)` and `workbookName(...)` can now also be chained
  after `workbook(...)`/`sheet(...)` on the returned source step, matching the export
  builders where `override(...)` may appear anywhere before the final step. Chaining
  them before the parse-target configuration keeps working; the value set last wins.
- `PxlSystemException`: the exception a builder's final (execute) step wraps a failure
  PXL does not classify in — a checked I/O failure, an unexpected runtime failure from
  POI, and so on — keeping the original accessible through `getCause()`.

### Changed

- **BREAKING** `PxlException` is now `abstract` and can no longer be instantiated
  (`new PxlException(...)`). Catching it and subclassing it are unaffected, and
  `throws PxlException` remains a valid contract on the final (execute) steps.
  What surfaces there is always a concrete subtype: the matching one for a classified
  failure, and `PxlSystemException` for everything else.
- Failing to open an Excel source (a missing file, an unreadable or password-protected
  container, an unsupported format) now surfaces as `PxlIOException` rather than the
  bare base type, so callers can catch that case on its own.

## [0.9.0] - 2026-07-24

First public release.

### Added

- Import: XLSX, XLS, and CSV into Java objects.
- Export: Java objects into XLSX (default), XLS, and streaming XLSX (SXSSF),
  selectable via `@PxlWorkbook(exportFileFormat = ...)`.
- Around 30 built-in field-type codecs (numbers, `BigInteger`/`BigDecimal`,
  full `java.time` including zoned/offset/`Duration`/`Period`, enums, collections,
  and custom objects), with per-column custom converters.

[Unreleased]: https://github.com/hclimkr/pxl/compare/v0.9.3...HEAD
[0.9.3]: https://github.com/hclimkr/pxl/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/hclimkr/pxl/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/hclimkr/pxl/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/hclimkr/pxl/releases/tag/v0.9.0
