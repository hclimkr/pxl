# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
  leaves no file behind and `toStream(...)` writes in one pass. The trade is memory proportional to output size —
  the 100,000-row cap does not bound it, as neither column count nor field length is capped in practice.
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
  the format; nothing narrows, so no CSV that imported before stops importing. The row cap stays at 100,000 —
  neither cap is a memory guard as implemented, both being checked after the file is already materialized.
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

[Unreleased]: https://github.com/hclimkr/pxl/compare/v0.9.3...HEAD
[0.9.3]: https://github.com/hclimkr/pxl/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/hclimkr/pxl/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/hclimkr/pxl/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/hclimkr/pxl/releases/tag/v0.9.0
