# PXL (Unreleased)

> Draft notes for the next release. The version is fixed when the release is cut — at that point this file is
> renamed to `vX.Y.Z.md`, its heading takes the version, and the matching `## [Unreleased]` section in
> `CHANGELOG.md` is retitled the same way.

CSV configuration release for **PXL** — the sheets of one CSV workbook can now be read with different encodings and delimiters, each sheet naming its own where it departs from the workbook, and option objects became immutable value holders built by their builder alone.

`@PxlSheet` gains two elements and `PxlImportSheetOption` two fields; no existing signature moved and no type was removed. Two things change for a caller — where `importCsvCharset` / `importCsvDelimiter` may be declared, and, the one breaking change, option objects becoming immutable: their setters and no-argument constructors are gone, so an option is built with its builder.

## Highlights

  - **A CSV workbook may mix encodings and delimiters across its sheets**: a CSV workbook is read as one file per sheet, which meant `importCsvCharset` / `importCsvDelimiter` sat one level too high — a workbook holding a legacy MS949 export alongside a UTF-8 one had no way to say so, and choosing either setting corrupted the other file. Both attributes now exist on `@PxlSheet` and on `PxlImportSheetOption` as well, resolving sheet option → `@PxlSheet` → workbook option → `@PxlWorkbook` → built-in default. A sheet that names neither inherits the workbook value, so nothing about an existing workbook changes:

    ```java
    @PxlWorkbook(importCsvCharset = "MS949")
    public class CompanyWorkbook {
        @PxlSheet(name = "Legacy")                                  // inherits MS949
        private List<CharsetRow> legacy;
        @PxlSheet(name = "Modern", importCsvCharset = "UTF-8")      // this one file only
        private List<CharsetRow> modern;
    }
    ```
  - **"Not specified" is a sentinel, so a sheet can name the default**: the annotation defaults became `""` and `'\0'` (`PxlConstants.UNSPECIFIED_IMPORT_CSV_*`) instead of `"UTF-8"` and `','`, which moved the effective defaults to the bottom of the cascade. That distinction is what makes `@PxlSheet(importCsvCharset = "UTF-8")` above mean something: were the annotation default a usable value, a sheet saying nothing would be indistinguishable from one naming UTF-8, the workbook value could never be reached, and `@PxlWorkbook(importCsvCharset = "MS949")` would have gone silently dead. The only behaviour a caller can observe is that an explicitly blank charset now falls back to `"UTF-8"` rather than reaching `Charset.forName("")` and throwing
  - **A null row in the data says so, instead of blaming a column**: exporting a collection with a `null` element failed with `PxlCellCodecException: sheet 'Employees', row 2, column 'Name': java.lang.NullPointerException` — the column named was simply the first one, no codec had run, and the message carried nothing but the exception's class name. It is now a `PxlDataException` naming the sheet and the position of the null element, raised before the row is created. The grouping branch had been inconsistent with the write loop about this, deliberately routing a null row to the `(ungrouped)` sheet and then failing to write it; the collection is checked once, up front, for both
  - **A CSV may carry 16,384 columns, not 100**: a header with more than 100 fields was rejected with `PxlDataException`, which meant the same row class could be exported to XLSX — 16,384 columns — and then fail to be read back from CSV. The old cap had no basis in the format, since CSV bounds neither rows nor columns; the new one matches the widest format PXL writes, so a file PXL produced stays a file PXL can read. It is still a ceiling, just one no real row class reaches. The row cap is unchanged at 100,000: that is the figure tied to memory, as a CSV is parsed into memory whole. Nothing narrows, so every CSV that imported before still imports
  - **Option objects are immutable and built only by their builder**: every field of the six `Pxl{Import,Export}{Workbook,Sheet,Column}Option` classes is `final` now, and `@Setter` / `@NoArgsConstructor` are gone, which drops 65 setters and 6 no-argument constructors from the public API. Nothing called them — inside the library, in its documentation, and across 470 tests, an option was always built with its builder, and the JavaBean pattern the setters existed for (`new X()` followed by `setX(...)`) appeared nowhere. The levels had drifted apart, too: the workbook options were already immutable while the sheet and column options were not. `@Getter`, `@AllArgsConstructor`, `@Builder` and the `add*Option(...)` methods stay, and a built option's child list is still a mutable `ArrayList`, so `option.addImportColumnOption(...)` works exactly as before
