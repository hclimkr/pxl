# PXL (Unreleased)

> Draft notes for the next release. The version is fixed when the release is cut — at that point this file is
> renamed to `vX.Y.Z.md`, its heading takes the version, and the matching `## [Unreleased]` section in
> `CHANGELOG.md` is retitled the same way.

CSV release for **PXL**: objects can now be written to CSV as well as read from it, and the sheets of one CSV workbook can carry their own encoding, delimiter and byte order mark in either direction. Two changes ask something of a caller — the import of `PxlFileFormat` / `PxlExcelEngine` moved, and option objects became immutable.

## Highlights

  - **Objects → CSV.** PXL has read XLS, XLSX and CSV from the start while writing only Excel. `exportCsv()` and `exportSampleCsv()` close that gap using the annotations, converters, column order, i18n column names and bean validation already in place — the CSV writer calls the same codec entry point in its cell-less form.
    ```java
    pxl.exportCsv()
       .sheet(Employee.class, employees, "Employees")
       .toFile(new File("employees.csv"));
    ```
    A CSV file holds one sheet, so the builder has the sheet form only and terminates at `toFile(File)` / `toStream(OutputStream)`. `exportSampleCsv()` writes the header-plus-one-sample-row template that `importCsv()` reads straight back as a filled-in form.
  - **What CSV cannot carry is dropped — with two exceptions and one refusal.** Stylers, widths, row heights, freeze panes, auto-filters, dropdowns, the engine and `exportGroupingFieldName` are ignored. Two attributes still put their value in the file: `exportStringAsFormula` writes the text as it stands, leading `=` and all, and `exportStringAsPicture` writes the image location — a path you expected to disappear is disclosed. `exportPassword` is refused instead of ignored, since CSV cannot be encrypted and writing plaintext would be a leak.
  - **Row and column coordinates are honored**, because CSV import already honors their `import*` counterparts. A row above the header, or a column before the first data column, is written as a record of empty fields rather than as a blank line — PXL's own import ignores blank lines, which would pull the header up on the way back in.
  - **Plan for the memory — and for the disk.** A CSV export renders its whole output before the destination is opened, which is what keeps a failure from leaving a file behind and lets `toStream(...)` write in one pass. Only the first 4 MiB of that output stays in memory: past `PxlConstants.EXPORT_MEMORY_THRESHOLD_OF_CSV` the rest continues into a temporary file under `java.io.tmpdir`, deleted before the call returns whether it succeeded or failed. So the heap a CSV export needs does not grow with the output — but a large one needs free disk space, and the temporary file is written in plaintext, which is worth knowing since a CSV export refuses `exportPassword` rather than encrypting.
  - **A CSV workbook may mix encodings, delimiters and byte order marks across its sheets — now in both directions.** One sheet is one file, so these settings sat one level too high: a workbook holding a legacy MS949 export alongside a UTF-8 one had no way to say so. `@PxlSheet` gains `importCsvCharset` / `importCsvDelimiter`, both annotations gain `exportCsvCharset` / `exportCsvDelimiter` / `exportCsvBom`, and the option classes follow — all resolving sheet option → `@PxlSheet` → workbook option → `@PxlWorkbook` → built-in default.
    ```java
    @PxlWorkbook(importCsvCharset = "MS949")
    public class CompanyWorkbook {
        @PxlSheet(name = "Legacy")                                  // inherits MS949
        private List<CharsetRow> legacy;
        @PxlSheet(name = "Modern", importCsvCharset = "UTF-8")      // this one file only
        private List<CharsetRow> modern;
    }
    ```
    The `@PxlWorkbook` import defaults became the sentinels `""` and `'\0'`, which is what lets the second field name UTF-8 explicitly; the only observable effect is that a blank charset now falls back to `"UTF-8"` instead of throwing. A byte order mark is written only for UTF-8, UTF-16LE and UTF-16BE — `UTF-16` has its encoder emit one already, and EUC-KR and the like cannot encode U+FEFF at all.
  - **`PxlOptionalBoolean`** (`UNSPECIFIED` · `TRUE` · `FALSE`) is the type `exportCsvBom` takes. An annotation element cannot hold `null`, so a `boolean` one has no value left to mean "not specified": its `false` would read as silence, and a sheet could never turn off a mark its workbook asked for.
  - **`PxlFileFormat` and `PxlExcelEngine` moved to `io.github.hclimkr.pxl.type`.** Update the import; nothing else about either type changed. They now sit with `PxlOptionalBoolean` in a package for the public enums the API takes as values, which keeps the base package to the entry point and the shared defaults.
  - **Option objects are immutable and built only by their builder.** Every field of the six `Pxl{Import,Export}{Workbook,Sheet,Column}Option` classes is `final`, and `@Setter` / `@NoArgsConstructor` are gone, dropping 65 setters and 6 no-argument constructors from the public API — nothing called them, as an option was always built with its builder. `@Getter`, `@AllArgsConstructor`, `@Builder` and the `add*Option(...)` methods stay.
  - **A null row in the data says so, instead of blaming a column.** Exporting a collection with a `null` element used to fail as a `PxlCellCodecException` blaming the first column — one no codec had touched — with `java.lang.NullPointerException` as its entire message. It is now a `PxlDataException` naming the sheet and the position of the null element, raised before the row is created, and the grouping branch no longer routes it to the `(ungrouped)` sheet only to fail there.
  - **A CSV may carry 16,384 columns, not 100.** A header with more than 100 fields was rejected, so the same row class could be exported to XLSX and then fail to be read back from CSV. The old cap had no basis in the format, since CSV bounds neither rows nor columns; the new one matches the widest format PXL writes. The row cap is unchanged at 100,000 — on import that is the figure tied to memory, since the file is read in full before the cap is checked — and nothing narrows, so every CSV that imported before still imports.
