# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/hclimkr/pxl/compare/v0.9.1...HEAD
[0.9.1]: https://github.com/hclimkr/pxl/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/hclimkr/pxl/releases/tag/v0.9.0
