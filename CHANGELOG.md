# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

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

### Fixed

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

[Unreleased]: https://github.com/hclimkr/pxl/compare/v0.9.2...HEAD
[0.9.2]: https://github.com/hclimkr/pxl/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/hclimkr/pxl/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/hclimkr/pxl/releases/tag/v0.9.0
