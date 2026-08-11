**English** · [한국어](reference_ko.md)

PXL Reference
=============================

[![Quick Start](https://img.shields.io/badge/🚀%20Quick%20Start-README.md-4c9aff?style=for-the-badge)](../README.md)

PXL is an **annotation-driven, bidirectional binding between spreadsheets and Java objects**.
It is built on top of Apache POI and Apache Commons CSV, and supports Java 8 and above.

- Internally, it handles Excel (XLS/XLSX) with Apache POI and CSV with Apache Commons CSV.
- Import: XLS, XLSX, CSV → Java objects
- Export: Java objects → Excel (XLSX by default; XLS and streaming XLSX can be selected via `@PxlWorkbook(exportExcelEngine = ...)`)
- Only annotated fields/classes become binding targets.

> 🚀 **New here → [README.md](../README.md)** — a hands-on guide that applies PXL the fastest, example-first.
> This document is the reference that covers supported variable types, the full set of options, and constraints.

## Table of Contents

1. [Why PXL](#why-pxl)
2. [Setup](#setup)
3. [API Structure](#api-structure)
4. [Supported Variable Types](#supported-variable-types)
5. [Annotations](#annotations)
6. [Option Override](#option-override)
7. [Validation](#validation)
8. [Exceptions](#exceptions)
9. [i18n](#i18n)
10. [Row/Column Index Rules Within a Sheet](#rowcolumn-index-rules-within-a-sheet)
11. [Cell Stylers](#cell-stylers)
12. [Various Examples](#various-examples)
13. [Limitation](#limitation)
14. [Common Pitfalls Checklist](#common-pitfalls-checklist)
15. [License](#license)

---

## Why PXL

Its goal is to handle the mapping between spreadsheets and Java objects as accurately and broadly as possible using only annotations.
With a declarative approach of putting annotations on your DTO, you get the following features without any extra code.

- **Type fidelity & strictness**   
  Full `java.time` (including `Zoned`/`Offset`/`Duration`/`Period`), `BigInteger`/`BigDecimal` precision (2^53) awareness,
  `NaN`/`Infinity` rejection, non-lenient date parsing (blocks rollover of invalid dates), `Collection` position preservation, symmetric import/export behavior.
  Not just common types — these edge cases are defended and documented per codec.
- **Standard validation integration + custom constraints**  
  Performs per-row validation on import with `javax.validation`/`jakarta.validation`,
  and separately provides custom constraints such as `@PxlByteSize` (byte length).
- **Header i18n**  
  Translates sheet/column names via `ResourceBundle` to build multilingual templates.
- **Simultaneous javax + jakarta support**  
  Generates both artifacts from a single source, covering both legacy (`javax`) and new (`jakarta`) environments.
- **Annotation-based sample/template export**  
  Generates an input form (including dropdowns, sample values, and i18n headers) directly from an annotated class.
- **A wide range of POI features via annotations**  
  Dropdowns/data validation, image insertion (including URLs), formulas (`exportStringAsFormula`), encryption,
  splitting sheets by field value, styler cascade, auto-size.
- **Multiple formats with a single annotation model**  
  Reading: XLS/XLSX/CSV, Writing: XLSX/XLS/streaming XLSX (SXSSF).

---

## Setup

Add only the variant that matches your environment to your dependencies.

**Maven**

```xml
<!-- javax variant (Java 8+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-javax</artifactId>
    <version>0.9.3</version>
</dependency>
```

```xml
<!-- jakarta variant (Java 17+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-jakarta</artifactId>
    <version>0.9.3</version>
</dependency>
```

**Gradle**

```groovy
// javax variant (Java 8+)
implementation 'io.github.hclimkr:pxl-javax:0.9.3'
```

```groovy
// jakarta variant (Java 17+)
implementation 'io.github.hclimkr:pxl-jakarta:0.9.3'
```

## Runtime Dependencies

The Bean Validation API (`validation-api` / `jakarta.validation-api`) and
the logging facade `slf4j-api` are declared by PXL with `compile` scope and thus transitively provided,
so you do not need to add them yourself.

You add the following two only when you need them.

### Bean Validation implementation + EL  

To perform data validation for `@NotNull`, `@NotEmpty`, `@NotBlank`, etc., you add an implementation (e.g., hibernate-validator) and EL (e.g., jakarta.el).  
If either one is missing, `new Pxl()` merely logs a warning (SLF4J `WARN`) and disables validation.  

**Maven example**

```xml
<!-- javax variant -->
<dependency>
    <groupId>org.hibernate.validator</groupId>
    <artifactId>hibernate-validator</artifactId>
    <version>6.2.5.Final</version>
</dependency>
<dependency>
    <groupId>org.glassfish</groupId>
    <artifactId>jakarta.el</artifactId>
    <version>3.0.4</version>
</dependency>
```

```xml
<!-- jakarta variant -->
<dependency>
    <groupId>org.hibernate.validator</groupId>
    <artifactId>hibernate-validator</artifactId>
    <version>9.1.1.Final</version>
</dependency>
<dependency>
    <groupId>org.glassfish</groupId>
    <artifactId>jakarta.el</artifactId>
    <version>4.0.2</version>
</dependency>
```

**Gradle example**

```groovy
// javax variant
implementation 'org.hibernate.validator:hibernate-validator:6.2.5.Final'
implementation 'org.glassfish:jakarta.el:3.0.4'
```

```groovy
// jakarta variant
implementation 'org.hibernate.validator:hibernate-validator:9.1.1.Final'
implementation 'org.glassfish:jakarta.el:4.0.2'
```

### SLF4J binding

To actually output logs, you add an SLF4J binding (`logback-classic`, `slf4j-simple`, `log4j-slf4j2-impl`, etc.).  
If no binding is present, SLF4J emits a `No SLF4J providers were found` warning and then discards logs (NOP).  
POI core logging uses `log4j-api`, which POI provides transitively, so no separate addition is needed.

**Maven example**

```xml
<dependency>
  <groupId>org.apache.logging.log4j</groupId>
  <artifactId>log4j-slf4j2-impl</artifactId>
  <version>2.26.1</version>
</dependency>
<dependency>
  <groupId>org.apache.logging.log4j</groupId>
  <artifactId>log4j-core</artifactId>
  <version>2.26.1</version>
</dependency>
```

**Gradle example**

```groovy
implementation 'org.apache.logging.log4j:log4j-slf4j2-impl:2.26.1'
implementation 'org.apache.logging.log4j:log4j-core:2.26.1'
```

---

## API Structure

### Method Chain

Each operation is performed as a single method chain, from the start method to the final method.  
The direction (export/import) and format (excel/csv) are embedded in the start method name.

| Purpose         | Method chain (start → configure → execute)                                                                                                                                                       |
|-----------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Excel export    | `pxl.exportExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                                     |
| Sample Excel export | `pxl.exportSampleExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                           |
| CSV export      | `pxl.exportCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                          |
| Sample CSV export | `pxl.exportSampleCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                  |
| Excel import    | `pxl.importExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromStream(InputStream)`                                                                                    |
| CSV import      | `pxl.importCsv()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromFiles(List<File>)` / `.fromStream(String, InputStream)` / `.fromStreams(List<String>, List<InputStream>)` |

- The configuration steps `.workbook(...)` and `.sheet(...)` are mutually exclusive — specifying both in one chain throws `PxlArgumentException` (as does omitting both).
- The two CSV export chains have no `.workbook(...)` at all: a CSV file holds one sheet, so there is no workbook form to call. `.sheet(...)` accumulates as it does for Excel, but the final method writes a single sheet, so configuring more than one throws `PxlArgumentException` there (`builder.export.csv.singleSheetOnly`) rather than at the configuration step.
- For export, calling `.sheet(...)` multiple times creates multiple sheets — the call order becomes the sheet order within the workbook, and each sheet may take a different row class. A duplicated sheet name (compared after normalization to a safe name, ignoring case) throws `PxlDataException` naming the offender — a workbook cannot hold two sheets whose names differ only in case.
- For import, `.sheet(...)` cannot be chained consecutively — there are two ways to read multiple sheets.
    - All at once, in workbook form: passing a `@PxlWorkbook` class to `.workbook(...)` binds multiple sheets at once, one per `@PxlSheet` field.  
      Since one CSV file is one sheet, in this form you pass multiple sources via `.fromFiles(List<File>)` / `.fromStreams(List<String>, List<InputStream>)` (the file name without its extension, or the `csvNames` entry, is matched against the `@PxlSheet` name — whitespace and case are ignored).
    - One sheet at a time: call `.sheet(...)` once per sheet on the same builder instance and run each through its final (execute) step. The source is given at the execute step, so open a fresh stream per call (files are opened and closed internally each time).
- The configuration-step `.override(...)` is optional, and its position within the chain can be set freely — before or after `.workbook(...)`/`.sheet(...)`, as long as it comes before the final (execute) step (if specified more than once, the last value wins). It overrides annotation values at runtime with the values carried in the option object.  
  For export, pass a `PxlExportWorkbookOption` to `.override(...)`; for import, pass a `PxlImportWorkbookOption` (if omitted, the annotation values are used as-is).
  For import, `.workbookName(String)`, which overrides the workbook name, can also be placed in the same position. Omit it and an Excel import from a file names the workbook after the file (extension removed); see [`@PxlWorkbookName`](#pxlworkbookname-targets-a-field).
- For the field list of each option and builder examples, see the [Option Override](#option-override) section.

### Resource Ownership

| Final method                                                                              | Parameter               | How PXL handles it                          |
|-------------------------------------------------------------------------------------------|-------------------------|---------------------------------------------|
| export<br/>`toFile(File)`                                                                 | Caller's `File`         | Opens the file's stream internally and closes it directly |
| export<br/>`toStream(OutputStream)`                                                       | Caller's `OutputStream` | Does not close it (flush only). The caller must close it   |
| export<br/>`toWorkbook()`                                                                 | Returned `Workbook`     | Does not close the returned workbook. The caller must close it |
| import<br/>`fromFile(File)`<br/>`fromFiles(List<File>)` (CSV)                             | Caller's `File`         | Opens the file's stream internally and closes it directly |
| import<br/>`fromStream(InputStream)` (Excel)<br/>`fromStream(String, InputStream)`·`fromStreams(List<String>, List<InputStream>)` (CSV) | Caller's `InputStream`  | Does not close it. The caller must close it |

A CSV export builds its whole output before the destination is touched, so `toFile(...)` creates no file when the
export fails and `toStream(...)` writes to the caller's stream once, in a single pass. A small output is held in
memory; a large one continues into a temporary file, which is deleted before the call returns; see
[Limitation](#limitation).

### Builder Lifecycle and Thread Safety

`Pxl` itself is stateless and thread-safe. The builders it hands out are the opposite: each carries the configuration collected so far, so they are not thread-safe.

| Object                                                  | Reuse                     | Note                                                                                             |
|---------------------------------------------------------|---------------------------|--------------------------------------------------------------------------------------------------|
| `Pxl`                                                   | Reuse it                  | Stateless and thread-safe; `new Pxl()` pays a validation-bootstrap cost, so keep it as a singleton |
| Export builders<br/>(`exportExcel()`·`exportSampleExcel()`) | Re-runnable, not re-configurable | The configuration stays on the builder: running a final (execute) method again repeats it, but adding more sheets accumulates on top of it |
| CSV export builders<br/>(`exportCsv()`·`exportSampleCsv()`) | Re-runnable, not re-configurable | Same as above, and `sheet(...)` accumulates the same way — but since a CSV terminal writes one sheet, a second sheet makes the terminal, not the configuration step, fail |
| Import builders<br/>(`importExcel()`·`importCsv()`)     | Reusable                  | The source step copies the settings it needs, so the same builder may be run once per sheet         |

An export builder keeps every sheet added to it, and no final (execute) method clears them.

That makes re-running the same configuration well-defined — each final (execute) method builds a fresh workbook out of the sheets held by the builder, so the same content can be sent to more than one destination:

```java
PxlExcelExportBuilder builder = pxl.exportExcel()
                                   .sheet(Employee.class, employees, "Employees");

builder.toFile(new File("employees.xlsx"));   // written
builder.toStream(outputStream);               // built again, same content
```

Note that the workbook really is built from scratch each time — this is a repeat, not a cached copy.

Re-configuring the same builder is what does not work. Calling `sheet(...)` again after a run adds to the sheets already there, so the next run writes the earlier ones as well, and a repeated sheet name raises `PxlDataException`. Start a fresh chain from `exportExcel()` for each workbook whose contents differ.

An import builder is the other way round. `workbook(...)`/`sheet(...)` hand their settings to the source step and leave the builder itself untouched, so re-configuring is exactly what it is for — drive as many sheets as you like off one builder:

```java
PxlExcelImportBuilder builder = pxl.importExcel();

List<Employee> employees = builder.sheet(Employee.class, "Employees")
                                  .fromFile(excelFile);
List<Department> departments = builder.sheet(Department.class, "Departments")
                                      .fromFile(excelFile);
```

Each call reads its own sheet, and the return type may differ per sheet. `workbookName(...)`/`override(...)` set on the builder apply to every run that follows. To apply one to a single run, set it on the source step instead — the step's value replaces the one copied from the builder rather than merging with it.

---

### Export Engine and File Format

Two enums sit on separate axes, and neither stands in for the other.

| Type              | Answers                                       | Constants                 |
|-------------------|-----------------------------------------------|---------------------------|
| `PxlExcelEngine`  | Which POI implementation writes the workbook   | `HSSF` · `XSSF` · `SXSSF` |
| `PxlFileFormat`   | What the resulting bytes are                   | `XLS` · `XLSX` · `CSV`    |

`@PxlWorkbook(exportExcelEngine = ...)` and `PxlExportWorkbookOption.exportExcelEngine` take the engine, so a format
no engine writes — CSV — cannot be declared there at all. Each engine knows the format it produces, and the
sheet/row/column limits belong to that format rather than to the engine, which is why `XSSF` and `SXSSF` share them.

Writing CSV is therefore chosen by the start method, not by an engine: `exportCsv()` / `exportSampleCsv()` take
a path of their own that produces no POI workbook, and the CSV limits apply because that path fixes the format to
`PxlFileFormat.CSV`. Anything an engine would have decided — the streaming window included — has no meaning there.

`PxlFileFormat` is the type to reach for when serving a download, as it carries the filename extension and the MIME
content type.

```java
PxlFileFormat format = PxlExcelEngine.SXSSF.getFileFormat();                 // XLSX

response.setContentType(format.getContentType());                            // application/vnd.openxmlformats-...
response.setHeader("Content-Disposition",
        "attachment; filename=report." + format.getFilenameExtension());     // report.xlsx
```

Both types can also be recovered from what you already hold, and every one of these lookups is plain: nothing is
thrown and `null` is never returned.

| Lookup                                       | Returns                                                                     |
|----------------------------------------------|-----------------------------------------------------------------------------|
| `PxlFileFormat.fromPoiWorkbook(workbook)`    | The format the workbook holds. A streaming-reader workbook answers `XLSX` like the other OOXML ones; `CSV` is never returned, as no POI workbook represents it |
| `PxlExcelEngine.fromPoiWorkbook(workbook)`   | The writer behind the workbook, telling `XSSF` and `SXSSF` apart            |
| `PxlExcelEngine.fromWorkbookObject(Class)`   | The engine a class declares through `@PxlWorkbook`                          |

---

## Supported Variable Types

| Category         | Types                                                                                        |
|------------------|----------------------------------------------------------------------------------------------|
| Primitive/wrapper | `byte` `short` `int` `long` `float` `double` `char` `boolean` and their wrapper classes, `String` |
| Numeric          | `BigInteger` `BigDecimal`                                                                    |
| Date/time        | `Date` `LocalDate` `LocalTime` `LocalDateTime` `ZonedDateTime` `OffsetTime` `OffsetDateTime` |
| Other            | `Enum`, user-defined classes, `Collection` of the above types                               |
| Experimental     | `Duration` `Period`                                                                          |

Fields of an unsupported variable type fail with `PxlArgumentException` while the column metadata is resolved (a user-defined class becomes a supported type once it has `@PxlImportConverter`/`@PxlExportConverter` or a `String` constructor). `PxlCellCodecException` is for a supported type whose cell value cannot be converted into it.

### Per-Type Behavior Summary — Import

| Type                                            | Import behavior / notes and limits                                                                                                                                                              |
|-------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `byte`·`short`·`int`·`long` + wrapper classes   | Numeric cell/string cell: parsed with `DecimalFormat` when a `pattern` is specified.<br/>Raises `PxlCellCodecException` when the type range is exceeded.<br/>Fractional part is truncated (e.g., `12.9`→`12`).<br/>For `long`/`Long`, a numeric cell is a double, so precision beyond 2^53 may be lost (accurate if a string cell) |
| `float`·`double` + wrapper classes              | Numeric cell/string cell: parsed at IEEE-754 precision.<br/>`NaN`·`Infinity`, and values beyond the `float` representable range, raise `PxlCellCodecException`.<br/>However, `float` underflow (e.g., `1e-300`→`0.0f`) is allowed as IEEE-754 precision loss |
| `char`·`Character`                              | String cell: the first character.<br/>Numeric cell: stringifies the integer/real as-is (`12`→`"12"`, `-3`→`"-3"`) and then takes only the first character. (`12`→`'1'`, `-3`→`'-'`).<br/>Boolean cell: `'1'`/`'0'`.<br/>Blank cell: `Character`=`null`, `char`=`' '` |
| `boolean`·`Boolean`                             | String cell: first compared case-insensitively against `importTrueString`/`importFalseString`, then against the built-in tokens `true/false`·`t/f`·`y/n`·`yes/no`·`on/off`·`1/0` (case-insensitive). If none match → exception (not a silent `false`).<br/>Numeric cell: `true` if not 0 |
| `String`                                        | String cell: the value as a string.<br/>Numeric/boolean cell: converted to a string                                                                                                            |
| `BigInteger`·`BigDecimal`                       | String cell: exactly restored with `new BigInteger/BigDecimal`.<br/>Numeric cell: limited to double (2^53) precision                                                                            |
| `Date`·`LocalDate`·`LocalTime`·`LocalDateTime`  | String cell: parsed with `pattern`/`importPattern` or the fixed ISO-8601 default pattern.<br/>Numeric cell: recognizes Excel date/time.<br/>Boolean cell: not supported, raises an exception.  |
| `ZonedDateTime`·`OffsetDateTime`·`OffsetTime`   | String cell: parsed with `pattern` or ISO-8601 including zone/offset, preserving the zone/offset.<br/>A string without an offset/zone raises an exception unless a `pattern` is given.<br/>Numeric cell: reads the Excel date/time based on the system default zone/offset.<br/>Boolean cell: raises an exception. |
| `Enum`                                          | Matched against the `toString()` override value or the constant name (ignoring case/whitespace).<br/>Custom conversion via `@PxlImportConverter`                                                |
| User-defined classes                            | Requires a single-argument `String` constructor or `@PxlImportConverter`                                                                                                                       |
| `Collection`                                    | Split by separator, preserving the positions of empty/null elements (e.g., `"a;;b"`→`["a", null, "b"]`).<br/>Elements must be concrete classes; nested generics (`List<List<..>>`), wildcards (`List<? extends X>`), and raw types raise `PxlReflectionException`. |
| `Duration`·`Period` (experimental)              | String cell: `pattern` or ISO-8601.<br/>Numeric cell: fixed unit (`Duration`=seconds, `Period`=days), fractional part truncated, raises an exception when the range is exceeded.                |

**Common to Import**

- Blank cells / empty values (Excel BLANK·missing cell, CSV empty value) do not set a value on the field — reference types become `null`, primitives get the DTO default (`0`/`false`, etc.).
- `importTrim` is `true` by default. When `false`, only `String` preserves whitespace; other types may fail to parse or get a wrong value if whitespace is mixed in.
- The Streaming Reader (`importUsingStreamReader`, XLSX-only) cannot evaluate formula cells and requires the header row position to be specified precisely.

### Per-Type Behavior Summary — Export

> The table describes an Excel export, where a cell has a type of its own. A CSV file has none, so what lands
> in the field is always the string the codec computed: a numeric or date value with no `pattern` is written as
> that codec's own text (`2023-06-15` for a `LocalDate`) rather than as a serial number carrying a display format,
> and the quote-prefix column below has no counterpart.

| Type                                            | Export behavior / notes and limits                                                                                                                                                                                        |
|-------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `byte`·`short`·`int` + wrapper classes          | Written as a numeric cell (safe because the representable range is under 2^53)                                                                                                                                            |
| `long`·`Long`                                   | Without a pattern, a numeric cell (double) → precision loss beyond 2^53.<br/>To preserve it, use a `pattern` or `BigInteger`/`BigDecimal`                                                                                  |
| `float`·`double` + wrapper classes              | Numeric cell                                                                                                                                                                                                             |
| `char`·`Character`                              | Written as a single-character string. `exportTrim` is not applied                                                                                                                                                        |
| `boolean`·`Boolean`                             | Written as a string according to `exportTrueString`/`exportFalseString`                                                                                                                                                  |
| `String`                                        | Writes text. Options: `exportTrim`, masking (`exportMasking`), `exportStringAsFormula` (leading `=` → formula), `exportStringAsPicture` (image)                                                                            |
| `BigInteger`·`BigDecimal`                       | Always written as a string cell to preserve precision → may be excluded from Excel sorting/formulas/filters.<br/>Can be rounded with `DecimalFormat` when a `pattern` is specified                                         |
| `Date`·`LocalDate`·`LocalTime`·`LocalDateTime`  | Without a `pattern`/`exportPattern` or masking, written as a numeric cell with a date format applied.<br/>The applied display format uses POI built-in format codes (locale-independent): `Date`·date=`m/d/yy`, time=`h:mm:ss`, datetime=`m/d/yy h:mm`.<br/>When a `pattern`/`exportPattern` or masking is specified, it is written as a string cell to preserve that value. |
| `ZonedDateTime`·`OffsetDateTime`·`OffsetTime`   | As above, written as a numeric cell without a pattern or masking.<br/>To preserve the zone/offset, output as a string with a `pattern` that includes the offset (e.g., `"yyyy-MM-dd HH:mm XXX"`)                           |
| `Enum`                                          | `toString()` override value or the constant name.<br/>Custom via `@PxlExportConverter`                                                                                                                                    |
| User-defined classes                            | Requires an overridden `toString()` or `@PxlExportConverter`                                                                                                                                                             |
| `Collection`                                    | Joined by separator, `null` element → empty slot (e.g., `["a", null, "b"]`→`"a;;b"`).<br/>`exportStringAsPicture` is possible                                                                                              |
| `Duration`·`Period` (experimental)              | Without a pattern, ISO-8601 (`toString()`).<br/>With a pattern, `DurationFormatUtils`; `Period` is an approximation because it is converted based on the current time                                                       |

**Common to Export**

- `null` values and empty/blank `String` values are written with the column's `exportNullString`, whose default is the empty string `""`.  
  That is, regardless of type, a `null` field is by default exported as a string cell containing an empty string.  
  You can specify a different string with `exportNullString`.
- Export generates XLSX by default; `exportExcelEngine` can also select the `HSSF` engine (XLS) or the `SXSSF` engine (streaming XLSX). CSV export is not supported — an engine is an Excel writer, so CSV cannot be named there.
- Sheet/column order is not guaranteed to follow the field declaration order, so if order matters, specify `exportOrder`.

---

## Annotations

Only annotated fields/classes are bound.
When you pass an option object to the `.override()` step to provide runtime values, those values override the annotation values.

### `@PxlWorkbook` (targets a class)

| Attribute                                                          | Default    | Description                                                                                                                                                                       |
|-------------------------------------------------------------------|------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `importPassword`                                                  | `""`       | Password to remove document protection on import                                                                                                                                  |
| `importDataValidation`                                            | `true`     | Whether to run Bean Validation (`@NotNull`·`@Size`, ... — ↓ *Validation*) over the imported result.<br/>Nothing to do with Excel's "data validation" feature; that is the dropdown written by `@PxlColumn(exportOptionItems)` (↓ *Column: Dropdown Export*) |
| `importUsingStreamReader`                                         | `false`    | Whether to use the Streaming Reader on import (XLSX only)                                                                                                                         |
| `importStreamReaderRowCacheSize`                                  | `100`      | Row cache size of the Streaming Reader                                                                                                                                            |
| `importStreamReaderBufferSize`                                    | `4096`     | Buffer size of the Streaming Reader                                                                                                                                               |
| `importCsvCharset`                                                | `""`→`"UTF-8"` | CSV only. Character encoding of the CSV to import, for every sheet of the workbook. A single sheet may depart from it with `@PxlSheet(importCsvCharset)`.<br/>Handles a leading BOM automatically (strips the UTF-8/UTF-16LE/BE BOM; for `UTF-16` (auto), the BOM is used to determine endianness) |
| `importCsvDelimiter`                                              | `'\0'`→`','` | CSV only. Delimiter of the CSV to import, for every sheet of the workbook (`char`). A single sheet may depart from it with `@PxlSheet(importCsvDelimiter)`                                   |
| `importI18nBaseName` / `importI18nLanguage` / `importI18nCountry` | `""`/`"en"`/`""` | Base name / language / country of the multilingual ResourceBundle on import                                                                                                       |
| `exportExcelEngine`                                               | `XSSF`     | POI engine that writes the workbook (`PxlExcelEngine`): `XSSF`=XLSX (default), `HSSF`=XLS, `SXSSF`=streaming XLSX.<br/>It selects the writer, not the format — `XSSF` and `SXSSF` both produce `.xlsx`. CSV is not an engine and cannot be named here. |
| `exportPassword`                                                  | `""`       | Document protection password to set on export.<br/>Applies to `toFile(...)`/`toStream(...)` only, not to `toWorkbook()`.<br/>A CSV export **rejects** it (`PxlArgumentException`) rather than writing plaintext. |
| `exportDataValidation`                                            | `true`     | Whether to run Bean Validation over the objects being written.<br/>Nothing to do with Excel's "data validation" feature — see `importDataValidation` above |
| `exportSXSSFRowAccessWindowSize`                                  | `100`      | rowAccessWindowSize on SXSSF export                                                                                                                                               |
| `exportCsvCharset`                                                | `""`→`"UTF-8"` | CSV only. Character encoding of the CSV to write, for every sheet of the workbook. A single sheet may depart from it with `@PxlSheet(exportCsvCharset)`                            |
| `exportCsvDelimiter`                                              | `'\0'`→`','` | CSV only. Delimiter of the CSV to write, for every sheet of the workbook (`char`). A single sheet may depart from it with `@PxlSheet(exportCsvDelimiter)`                         |
| `exportCsvBom`                                                    | `UNSPECIFIED`→`false` | CSV only. Whether a byte order mark precedes the output (`PxlOptionalBoolean`). A single sheet may depart from it with `@PxlSheet(exportCsvBom)`.<br/>Honored for UTF-8/UTF-16LE/UTF-16BE only; any other charset drops it silently (↓ *Limitation*) |
| `exportWorkbookRequiredHeaderCellStyler`                          | (unspecified) | Required header cell style (uses `PxlHeaderRequiredStyler` when unspecified/not applicable)                                                                                       |
| `exportWorkbookOptionalHeaderCellStyler`                          | (unspecified) | Optional header cell style (uses `PxlHeaderOptionalStyler` when unspecified/not applicable)                                                                                       |
| `exportWorkbookDataCellStyler`                                    | (unspecified) | Data cell style (uses `PxlDataVerticalCenterTextStyler` when unspecified/not applicable)                                                                                          |
| `exportI18nBaseName` / `exportI18nLanguage` / `exportI18nCountry` | `""`/`"en"`/`""` | Base name / language / country of the multilingual ResourceBundle on export                                                                                                       |

### `@PxlWorkbookName` (targets a field)

Put it on the `String` field that holds the workbook name; it can be omitted if not needed.  
Putting it on a field that is not a `String` raises `PxlDataException`.  
Import-only: on import in the workbook form, the field is filled with the name passed to the builder's `.workbookName(...)`.
When no name is given and the source is read from a file, it is set to that file's name with the extension removed.
It is not used on export (or sample export).

### `@PxlSheet` (targets a field)

Put it on a `Collection`-type field to bind it as a sheet.  
The default value `0` of an index attribute means "auto" (first row/column automatic).

| Attribute                                                                                                    | Default | Description                                              |
|-------------------------------------------------------------------------------------------------------------|---------|---------------------------------------------------------|
| `name`                                                                                                      | field name | Sheet name (array). Must match the actual sheet name to bind (whitespace and case ignored).<br/>When specified as an array, only one of them must exist |
| `importEnabled`                                                                                             | `true`  | Whether import is enabled                               |
| `importOverrideSuperClassSheet`                                                                             | `false` | Whether to override the superclass field of the same sheet name (case ignored) |
| `importExcludeHiddenRows` / `importExcludeHiddenColumns`                                                    | `false` | Whether to exclude hidden rows/columns                  |
| `importEachCellOfMergedRegion`                                                                              | `false` | Whether to treat a merged cell as the same value in each individual cell |
| `importHeaderRowIndex` / `importFirstDataRowIndex` / `importLastDataRowIndex`                               | `0`     | Header/first/last data row on import (1-based, see *Index Rules* below) |
| `importFirstDataColumnIndex` / `importLastDataColumnIndex`                                                  | `0`     | First/last data column on import (1-based)              |
| `importCsvCharset`                                                                                          | (inherit) | CSV only. Character encoding used to read this sheet's CSV. Blank (`""`) inherits the workbook value |
| `importCsvDelimiter`                                                                                        | (inherit) | CSV only. Delimiter used to read this sheet's CSV (`char`). NUL (`'\0'`) inherits the workbook value |
| `exportEnabled` / `exportSampleEnabled`                                                                     | `true`  | Whether export / sample export is enabled               |
| `exportOverrideSuperClassSheet`                                                                             | `false` | Whether to override the superclass field of the same sheet name (case ignored) |
| `exportRowHeightInPoints`                                                                                   | `-1.0`  | Row height within the sheet (points). Default height if unset |
| `exportOrder`                                                                                               | `""`    | Sheet creation order key (string comparison, see *Export Order* below) |
| `exportGroupingFieldName`                                                                                   | `""`    | Group by this field value and split into multiple sheets |
| `exportHeaderRowIndex` / `exportFirstDataRowIndex` / `exportLastDataRowIndex`                               | `0`     | Header/first/last data row on export (1-based)          |
| `exportFirstDataColumnIndex` / `exportLastDataColumnIndex`                                                  | `0`     | First/last data column on export (1-based)              |
| `exportIfNull`                                                                                              | `false` | Whether to create the sheet when the field is null      |
| `exportIfEmpty`                                                                                             | `true`  | Whether to create the sheet when the field is empty     |
| `exportColumnFilter`                                                                                        | `false` | Whether to apply a filter                               |
| `exportCsvCharset`                                                                                          | (inherit) | CSV only. Character encoding used to write this sheet's CSV. Blank (`""`) inherits the workbook value |
| `exportCsvDelimiter`                                                                                        | (inherit) | CSV only. Delimiter used to write this sheet's CSV (`char`). NUL (`'\0'`) inherits the workbook value |
| `exportCsvBom`                                                                                              | (inherit) | CSV only. Whether a byte order mark precedes this sheet's CSV (`PxlOptionalBoolean`). `UNSPECIFIED` inherits the workbook value, and `FALSE` turns off a mark the workbook asked for — which a plain `boolean` could not express |
| `exportSheetRequiredHeaderCellStyler` / `exportSheetOptionalHeaderCellStyler` / `exportSheetDataCellStyler` | (unspecified) | Sheet-level cell style (delegates to the Workbook level when unspecified) |

### `@PxlRowIndex` (targets a field)

Put it on a field that holds the row index; it can be omitted if not needed.  
The field type supports `byte`·`short`·`int`·`long` and their wrapper classes (`Byte`·`Short`·`Integer`·`Long`); any other type fails with `PxlArgumentException`.  
The value filled in is the 1-based spreadsheet row number of the imported row (the same numbering shown in the spreadsheet UI and used by `importHeaderRowIndex` and the other index attributes).  
In the default configuration (header on the first row), data rows are 2, 3, 4…, and if you move the header down with `importHeaderRowIndex` or there is a title row above the header, the absolute row number grows accordingly.  
Import-only: It is not used on export (or sample export).

### `@PxlColumn` (targets a field)

| Attribute                                                                                                       | Default    | Description                                                                                                     |
|----------------------------------------------------------------------------------------------------------------|------------|----------------------------------------------------------------------------------------------------------------|
| `name`                                                                                                         | field name | Column name (array). Must match the actual column name to bind (whitespace ignored, case-sensitive).<br/>When specified as an array, only one of them must exist |
| `pattern`                                                                                                      | `""`       | Common fallback pattern used when `importPattern` / `exportPattern` are empty                                   |
| `collectionSeparator`                                                                                          | `";"`      | Common fallback used when `importCollectionSeparator` / `exportCollectionSeparator` are empty                   |
| `importEnabled`                                                                                                | `true`     | Whether import is enabled                                                                                       |
| `importTrim`                                                                                                   | `true`     | Whether to trim the string on import.<br/>When `false`, numbers/dates/`Boolean`, etc., may fail to parse or get a wrong value due to whitespace |
| `importUnique`                                                                                                 | `false`    | Whether to check uniqueness of column values on import                                                          |
| `importPattern`                                                                                                | `""`       | Import format (numeric=`DecimalFormat`, date/time=`DateTimeFormat`).<br/>Date/time falls back to the default pattern on failure |
| `importTrueString` / `importFalseString`                                                                       | `"true"`/`"false"` | `String` column: renders a boolean cell as this string.<br/>`Boolean` column: interprets this string (case-insensitive) as true/false (takes priority over the built-in tokens) |
| `importCollectionSeparator`                                                                                    | `""`       | Separator to split a cell value into Collection elements.<br/>The entire literal string (`"::"`·`", "`, etc., multi-character allowed). |
| `importOverrideSuperClassColumn`                                                                               | `false`    | Whether to override the superclass field of the same column name                                               |
| `exportEnabled` / `exportSampleEnabled`                                                                        | `true`     | Whether export / sample export is enabled                                                                       |
| `exportSample`                                                                                                 | `""`       | Value to put in the cell on sample export.<br/>Parsed into the column type, so it must be a value that type accepts (`PxlCellCodecException` otherwise).<br/>Left unset, the sample cell gets `exportNullString` |
| `exportTrim`                                                                                                   | `false`    | Whether to trim the string on export (not applied to `char`/`Character` columns)                                |
| `exportPattern`                                                                                                | `""`       | Export format (numeric=`DecimalFormat`, date/time=`DateTimeFormat`, `Duration`/`Period`=`DurationFormatUtils`). |
| `exportColumnWidth`                                                                                            | `0`(auto)  | Column width. The default (0=`autoSizeColumn`) measures all rows, so it causes **performance degradation on large-data export** (↓ *Limitation* section).<br/>If there are many rows, a fixed width is recommended |
| `exportCollectionSeparator`                                                                                    | `""`       | Collection element separator string.                                                                           |
| `exportOverrideSuperClassColumn`                                                                               | `false`    | Whether to override the superclass field of the same column name                                               |
| `exportOrder`                                                                                                  | `""`       | Column creation order key (string comparison, see *Export Order* below)                                         |
| `exportMasking`                                                                                                | `""`       | Regex for the part to mask.<br/>Not applied to `char`/`Character` columns                                       |
| `exportOptionItems`                                                                                            | `{}`       | List of selectable options (dropdown)                                                                           |
| `exportEnumDropDownListStyle`                                                                                  | `SET`      | Style to set an Enum field as a dropdown (`SET` / `SORTED_SET` / `NONE`)                                        |
| `exportNullString`                                                                                             | `""`       | String to use when exporting a null value (and empty/blank `String`). Default is a string cell containing an empty string (not a blank cell) |
| `exportTrueString` / `exportFalseString`                                                                       | `"true"`/`"false"` | Strings to use when exporting true/false.<br/>To import a custom value again, specify `importTrueString`/`importFalseString` with the same values |
| `exportStringAsPicture`                                                                                        | `false`    | Insert an image URL string into the cell as an image                                                            |
| `exportStringAsFormula`                                                                                        | `false`    | Compute a formula string (leading `=`) and apply it to the cell                                                 |
| `exportColumnRequiredHeaderCellStyler` / `exportColumnOptionalHeaderCellStyler` / `exportColumnDataCellStyler` | (unspecified) | Column-level cell style (delegates to the Sheet level when unspecified)                                      |

### `@PxlImportConverter` / `@PxlExportConverter` (targets a method)

Specify custom String ↔ object conversion for an Enum or user-defined class.

```java
// String → value (static, return type is the target type)
@PxlImportConverter
public static EnumOrObjectType pxlImportConverter(final String str) {
    return ...;
}

// value → String
@PxlExportConverter
public String pxlExportConverter() {
    return ...;
}
```

> **Constraints:**   
> A `@PxlImportConverter` method must be `static`, and its return type must match the target type (enum/user-defined class) (since it is a factory that creates a new
> value from a string, instance methods are not supported).  
> `@PxlExportConverter` can be either an instance method (`String pxlExportConverter()`) or a `static` method that takes the target value as an argument
> (`static String pxlExportConverter(Type value)`), and its return type must be `String`.

---

## Option Override

### Option Usage

Option objects (`PxlImportWorkbookOption`·`PxlExportWorkbookOption`) are a bundle of values that override, at runtime, the values declared with the `@PxlWorkbook`·`@PxlSheet`·`@PxlColumn` annotations, created per workbook.  
When you pass this object to the `.override(...)` configuration step before the final (execute) step, the carried values overwrite the annotation values — fields not specified follow the annotation values (or defaults if none).  
The `.override(...)` step itself can also be omitted, in which case the annotation values are used as-is.

```java
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;

// 1) Build the option object with the builder (per workbook)
PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
        .exportPassword("secret")          // encrypt the document
        .exportDataValidation(false)       // turn off validation of export data
        .build();

// 2) Pass it via .override(...) before execution to override the annotation values
pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .override(exportOption)                                   // overrides @PxlWorkbook(exportPassword=...) etc.
   .toFile(new File("secured.xlsx"));
```

```java
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

// 1) Build the option object with the builder (per workbook)
PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
        .importDataValidation(false)       // turn off validation of import results (default true)
        .importUsingStreamReader(true)     // stream a large XLSX
        .build();

// 2) Pass it via .override(...) before execution to override the annotation values
List<Employee> rows = pxl.importExcel()
                         .override(importOption)             // overrides the annotation values
                         .sheet(Employee.class, "Employees")
                         .fromFile(file);
```

### Option Structure

Options are not a single workbook level but a 3-level tree of workbook → sheet → column, so they can be passed with a single `.override(...)`.

- Put sheet options (`Pxl{Import,Export}SheetOption`) in the workbook option's `importSheetOptions`/`exportSheetOptions`.
- Put column options (`Pxl{Import,Export}ColumnOption`) in the sheet option's `importColumnOptions`/`exportColumnOptions`.
- Attach child options with the builder's list setter (`.importColumnOptions(List)`, etc.) or the `add*Option(...)` method.
- Matching key: a sheet option is linked to its target by `fieldName` (the `@PxlSheet` field name of the workbook class), and a column option by `fieldName` (the `@PxlColumn` field name of the row class).  
  If you omit a sheet option's `fieldName`, it applies to all sheets as a wildcard (`*`); the single-sheet form (`sheet(...)`) uses this approach.
- Renaming at runtime: a sheet option overrides `@PxlSheet(name)` with `importSheetNames`/`exportSheetNames`, and a column option overrides `@PxlColumn(name)` with `importColumnNames`/`exportColumnNames` (a list, so aliases work the same as in the annotation).  
  They replace the annotation `name`, so they are bundle keys as well — with i18n on, an overriding name is translated before it is matched or written (↓ [i18n](#i18n)).
- Any level/field not specified follows the annotation value (or default if none).

```java
// Workbook → sheet (wildcard) → column (age) 3-level override: exclude only the age column from binding
PxlImportColumnOption ageColumn = PxlImportColumnOption.builder()
        .fieldName("age")                                   // match by column (@PxlColumn field name)
        .importEnabled(false)                               // do not bind this column
        .build();
PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()   // fieldName omitted → wildcard (all sheets)
        .importColumnOptions(Arrays.asList(ageColumn))
        .build();
PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
        .importSheetOptions(Arrays.asList(sheetOption))
        .build();

List<Employee> rows = pxl.importExcel()
                         .override(option)                          // pass the entire tree at once
                         .sheet(Employee.class, "People")
                         .fromFile(file);
// → age is not bound, so it keeps its initial value (null for Integer); the other columns bind normally
```

Export is configured the same way with `exportSheetOptions`/`exportColumnOptions`.

---

## Validation

### Standard Constraint

If you put validation annotations (`@NotNull`, `@NotEmpty`, `@NotBlank`, `@Valid`, ...) on fields, the import result object is validated,
and on failure a `PxlValidationException` is raised.  
Nested objects (sheet lists) propagate via `@Valid`.

```java
@NotBlank
@PxlColumn(name = "Name")
private String name;

// Nested lists of the workbook/sheet propagate via @Valid
@NotEmpty
@Valid
@PxlSheet(name = "Employees")
private List<Employee> employees;
```

Validation is controlled by `@PxlWorkbook(importDataValidation = ...)` (default `true`); turning it off means validation is not performed.
However, this attribute only concerns whether a field's value satisfies the constraints.  
If a column marked required with `@NotNull`·`@NotEmpty`·`@NotBlank` is missing from the sheet header, an exception is raised even with validation off — because whether a required column actually exists is always checked separately from value validation.  
If a Bean Validation implementation/EL is not on the classpath, validation is silently disabled (see [Setup](#setup)).

### Custom Constraint — `@PxlByteSize`

Along with standard validation constraints (`@NotNull`·`@Size`, etc.), PXL provides `@PxlByteSize`, which validates the byte length of a string.  
Unlike `@Size`, which counts characters, `@PxlByteSize` validates the number of bytes encoded with a specified charset.  
It is useful for directly expressing a column constraint where one Korean character takes multiple bytes, like a DB `VARCHAR(n BYTE)`.

```java
import io.github.hclimkr.pxl.constraint.PxlByteSize;

// Max 30 bytes in UTF-8 (10 Korean characters = 30 bytes)
@PxlByteSize(max = 30)                              // charset default "UTF-8"
@PxlColumn(name = "Name")
private String name;

// Range + explicit charset (EUC-KR: 1 Korean character = 2 bytes)
@PxlByteSize(min = 4, max = 20, charset = "EUC-KR")
@PxlColumn(name = "Code")
private String code;
```

- The target type is `CharSequence` (mainly `String`). `null` is considered valid.
- `min`/`max` are byte lengths (bounds inclusive). `charset` defaults to `"UTF-8"` when unspecified.
- When `importDataValidation` is on, the import result object is validated, and on failure a `PxlValidationException` is raised.

---

## Exceptions

### Exception Types

Exceptions that cross the boundary are all normalized at the `Pxl` boundary into the checked `PxlException` family.  
Exception message text is localized into multiple languages.  
The default is English, Korean is provided, and the language is set globally with `Pxl.setMessageLocale(Locale)`.  
For details, see "Exception/Diagnostic Message Language" in [i18n](#i18n).

| Exception                 | When it occurs                          |
|---------------------------|-----------------------------------------|
| `PxlException`            | abstract base type — never thrown itself; the common supertype of every checked Pxl exception |
| `PxlCellCodecException`   | When a cell value cannot be converted to the target type |
| `PxlValidationException`  | On validation failure                   |
| `PxlReflectionException`  | On reflection failure                   |
| `PxlArgumentException`    | On an argument/annotation configuration error |
| `PxlNullPointerException` | When a required (non-null) argument is `null` |
| `PxlDataException`        | When the workbook/CSV shape does not match what the binding expects — a target sheet/column that cannot be found or is duplicated, a sheet/row/column count limit exceeded, a missing header row/column, nothing to export, a `null` element in the collection to export, counts that do not line up (names vs. classes vs. streams), or invalid row/column index settings |
| `PxlIOException`          | On an I/O error — a source that cannot be opened or read (a missing file, an unsupported format, or decryption failing on a wrong `importPassword`), or a failed write |
| `PxlI18nException`        | When an i18n ResourceBundle cannot be found |
| `PxlSystemException`      | On a failure PXL does not classify — anything not covered by the types above (e.g. an unexpected runtime failure from POI), wrapped at the boundary with the original as its `getCause()` |
| `PxlRuntimeException`     | unchecked exception — currently unused  |

All checked exceptions such as `PxlNullPointerException`·`PxlValidationException` are subtypes of `PxlException`.  
`PxlException` itself is abstract, so what you actually catch is always a concrete subtype — the matching one for a classified failure, and `PxlSystemException` for everything else.  
A `throws PxlException` declaration on a final (execute) step therefore remains a valid contract, and callers may either handle them together with `catch (PxlException e)` or narrow to whichever subtypes they can act on.  
However, `PxlRuntimeException` is the exception: it is unchecked (in the `RuntimeException` family), is not a subtype of `PxlException`, and is currently unused.  

### Exception/Diagnostic Message Language

The exception messages and diagnostic log text that PXL throws support multiple languages.  
This text is resolved from the bundle (`pxl-messages`) that the library bundles into its artifact; the default is English, and Korean is provided.

- Determined by the process-wide locale.  
  The default is the JVM default locale (`Locale.getDefault()`), set globally with `Pxl.setMessageLocale(Locale)` and cleared with `Pxl.resetMessageLocale()`.  
  If there is no matching translation, it falls back to English.
- It is independent of content i18n (sheet/column names, `@PxlWorkbook`-based·per-workbook). You can produce English output while viewing server-log exceptions in Korean, or vice versa.

```java
Pxl.setMessageLocale(Locale.ENGLISH);   // subsequent exception/diagnostic text in English
// ...
Pxl.resetMessageLocale();               // revert to the JVM default locale
```

---

## i18n

If you specify a `ResourceBundle` via `@PxlWorkbook`'s `import/exportI18nBaseName`, `import/exportI18nLanguage`, `import/exportI18nCountry`,
sheet/column names are translated for matching/output (UTF-8 properties supported).  
The `name` value of `@PxlColumn`/`@PxlSheet` becomes the bundle key. It is an ordinary `ResourceBundle` key, and the
bundle is usually shared with the rest of the application, so namespace it (`staff.column.role`) instead of using a
bare word that another message could collide with.

i18n is disabled by default (opt-in).  
It works only when `import/exportI18nBaseName` is explicitly specified (or a `ResourceBundle` is injected into the option); if the base name is empty, no bundle is loaded and the name is used as-is.  
If a base name is specified but its `ResourceBundle` cannot be found, it fails with `PxlI18nException`.

Instead of naming a base name, you can hand PXL a bundle you already hold, through the workbook option's
`importResourceBundle`/`exportResourceBundle`.  
An injected bundle takes precedence over the annotation: the
`import/exportI18nBaseName`·`Language`·`Country` triple is not even loaded, so it cannot fail with `PxlI18nException`.  
Use it when the bundle comes from somewhere the annotation cannot name — a container-managed `MessageSource`, or a
locale chosen per request.

```java
ResourceBundle bundle = ResourceBundle.getBundle("messages", userLocale);   // resolved by the application

List<Person> rows = pxl.importExcel()
                       .override(PxlImportWorkbookOption.builder()
                                                        .importResourceBundle(bundle)   // wins over @PxlWorkbook(importI18nBaseName)
                                                        .build())
                       .sheet(Person.class, "staff.sheet")
                       .fromFile(file);
```

`src/main/resources/messages.properties` — the base bundle, read as UTF-8:

```properties
staff.sheet=Staff
staff.column.role=Role
staff.column.fullName=Full Name
```

`src/main/resources/messages_ko.properties` — the variant `exportI18nLanguage = "ko"` picks instead:

```properties
staff.sheet=직원
staff.column.role=역할
staff.column.fullName=성명
```

```java
@PxlWorkbook(
        exportI18nBaseName = "messages", exportI18nLanguage = "en",
        importI18nBaseName = "messages", importI18nLanguage = "en")
public class StaffWorkbook {

    @PxlSheet(name = "staff.sheet")                 // header/sheet name is translated to "Staff"
    private List<Person> people;
}

public class Person {

    @PxlColumn(name = "staff.column.role")          // header is translated to "Role"
    private String role;

    @PxlColumn(name = "staff.column.fullName")      // header is translated to "Full Name"
    private String fullName;
}
```

Besides names, two `@PxlColumn` attributes go through the same bundle on export:

- **`exportSample`** on a `String`/enum column. When the column is a `Collection` of those, the sample holds one key per element, split by `exportCollectionSeparator`, and **each element is translated on its own** — the separator stays in the annotation instead of being baked into a bundle value. An enum sample is parsed back into its constant after translation, so the cell ends up holding the canonical name.
- **`exportOptionItems`** on a `String` column, so the dropdown offers the very text the cells hold. A column of any other type writes its value in canonical form, so its items — and the enum constants used when no items are given — are taken verbatim; translating them would leave the written value outside the list Excel validates it against.

`messages.properties`:

```properties
staff.column.role=Role
staff.column.roles=Roles
staff.role.admin=Administrator
staff.role.user=User
```

```java
@PxlColumn(name = "staff.column.role", exportSample = "staff.role.admin",
        exportOptionItems = {"staff.role.admin", "staff.role.user"})   // sample cell "Administrator", dropdown Administrator/User
private String role;

@PxlColumn(name = "staff.column.roles", exportSample = "staff.role.admin;staff.role.user")   // sample cell "Administrator;User"
private List<String> roles;
```

> It works only when `i18nBaseName` is specified (untranslated by default). If specified but the bundle cannot be found, it does not silently pass but raises `PxlI18nException`.

---

## Row/Column Index Rules Within a Sheet

Index attributes of `@PxlSheet` such as `importHeaderRowIndex`, `exportFirstDataColumnIndex` are all 1-based,
and the default value `0` means auto (first/last automatic).

| Attribute              | Default   | Constraint                                       |
|------------------------|-----------|--------------------------------------------------|
| `HeaderRowIndex`       | first row | Must be smaller than `FirstDataRowIndex`         |
| `FirstDataRowIndex`    | second row | Greater than `HeaderRowIndex` and at most `LastDataRowIndex` |
| `LastDataRowIndex`     | last row  | At least `FirstDataRowIndex`                      |
| `FirstDataColumnIndex` | first column | At most `LastDataColumnIndex`                  |
| `LastDataColumnIndex`  | last column | At least `FirstDataColumnIndex`                 |

> Validation on export:  
> - If the column range specified by `exportLastDataColumnIndex` is smaller than the number of columns to export (some columns missing), an exception is raised.
> - If there are duplicate column names within the same sheet, an exception is raised.

> A CSV export honors all five the same way an Excel export does, since the coordinates are resolved in the shared
> metadata. What a coordinate looks like in a text file is worth knowing, though: a row that precedes the header,
> or a column before `FirstDataColumnIndex`, is written as an empty field rather than as a blank line — a
> record of empty fields (`"",,,`) instead of nothing at all. That is deliberate. PXL's CSV import ignores blank
> lines, so writing them would pull the header up on the way back in and break the round trip.
  
> Empty data range behavior on import:  
> - With a configuration where the data row range does not overlap the actual data or becomes empty (e.g., specifying `importFirstDataRowIndex` at a row larger than the actual data),
>   that sheet is processed as an empty result (empty collection) without error (no exception if the header row/required columns are valid).
> - However, a direct inversion where `importLastDataRowIndex` is explicitly specified smaller than `importFirstDataRowIndex` is raised as a `PxlDataException`.

---

## Cell Stylers

Cell style is delegated in the order column → sheet → workbook → built-in default.
At each stage, if a styler is unspecified or cannot be applied, it descends to the next stage.

| Purpose       | Default Styler                     |
|---------------|------------------------------------|
| Required header | `PxlHeaderRequiredStyler`        |
| Optional header | `PxlHeaderOptionalStyler`        |
| Data          | `PxlDataVerticalCenterTextStyler`  |

- Built-in header stylers: `PxlHeaderRequiredStyler` · `PxlHeaderOptionalStyler` · `PxlHeaderHorizontalCenterTextStyler` ·
  `PxlHeaderVerticalCenterTextStyler` · `PxlHeaderWrapTextStyler`.
- Built-in data stylers: `PxlDataTextStyler` · `PxlDataThinBorderStyler` · `PxlDataVerticalCenterTextStyler` ·
  `PxlDataHorizontalCenterTextStyler` · `PxlDataWrapTextStyler` · `PxlDataCommaSeparatedNumericStyler`.
- To implement your own, implement `PxlStyler` (`Font apply(Workbook, CellStyle)`).  
  Modify the `CellStyle` and return the `Font` to apply to the cell. (For an example, see [Column: Cell Styler Export](#column-cell-styler-export)).

---

## Various Examples

### Column: Date and Number Format

```java
@PxlColumn(name = "HireDate", pattern = "yyyy/MM/dd")
private LocalDate hireDate;

@PxlColumn(name = "StartTime", pattern = "HH:mm")
private LocalTime startTime;

@PxlColumn(name = "Amount", pattern = "#,##0.00")   // numeric uses DecimalFormat
private BigDecimal amount;
```

> To use different formats for import/export, provide `importPattern` / `exportPattern` separately.  
> For date/time, the format is not strictly enforced on import (parsed with the fixed ISO default pattern) but is enforced on export.  
> The default write pattern is ISO-8601, so the output of a value without a pattern is the same on any machine.  
> The default read pattern of LocalDateTime accepts both `T` (`yyyy-MM-dd'T'HH:mm:ss`) and space (`yyyy-MM-dd HH:mm:ss`). 

### Column: True/False Strings

```java
@PxlColumn(name = "Flag",
        exportTrueString = "Y", exportFalseString = "N",
        importTrueString = "Y", importFalseString = "N")   // use the same values for import/export to round-trip
private Boolean flag;
```

### Column: Collection

```java
@PxlColumn(name = "Tags", collectionSeparator = ",")   // the default separator is ";"
private List<String> tags;   // "a,b,c" ↔ [a, b, c]
```

> The positions (indices) of empty/`null` elements are preserved (`"a;;b"` ↔ `["a", null, "b"]`).  
> The element type must be a concrete class; nested generics (`List<List<..>>`) and wildcards are not supported.

### Column: Enum

```java
@PxlColumn(name = "Grade")
private Grade grade;   // matched against the toString() override value or the constant name (ignoring case/whitespace)
```

### Column: Custom Object

Put `@PxlImportConverter` (static, return type = target type) / `@PxlExportConverter` (instance or static, returns `String`).

```java
@Getter                     // (optional) for your convenience — PXL converts this type through the converters below, so a getter is not required.
@Setter                     // (optional) for your convenience — PXL converts this type through the converters below, so a setter is not required.
@NoArgsConstructor          // (optional) PXL never instantiates a custom object with a no-arg constructor.
@AllArgsConstructor         // used by parse() below — PXL itself does not require it.
public class Money {

    private String currency;
    private long amount;

    @PxlImportConverter                       // String → Money (must be static)
    public static Money parse(final String value) {
        final String[] p = value.trim().split("\\s+");
        return new Money(p[0], Long.parseLong(p[1]));
    }

    @PxlExportConverter                        // Money → String
    public String toCell() {
        return currency + " " + amount;
    }
}
```

```java
@PxlColumn(name = "Price")
private Money price;
```

> Without a converter, import uses the single-argument `String` constructor and export uses the overridden `toString()`.

### Column: Masking Export

```java
// Export with characters matching the regex replaced by '*' (the masked value round-trips as-is)
@PxlColumn(name = "Secret", exportMasking = "\\d")     // mask all digits
private String secret;
```

### Column: Dropdown Export

```java
// Fixed-list dropdown
@PxlColumn(name = "Choice", exportOptionItems = {"Red", "Green", "Blue"})
private String choice;
```

### Column: null Export

```java
// Export null as a specific string (the default is the empty string "")
@PxlColumn(name = "Memo", exportNullString = "-")
private String memo;
```

### Column: Formula Export

```java
// Compute a string with a leading '=' as a formula and apply it to the cell (the computed result is read on non-streaming import)
@PxlColumn(name = "Total", exportStringAsFormula = true)
private String total;   // e.g., "=A2+B2"
```

### Column: Image Export

```java
// Insert a string (image URL/path) as an actual image
@PxlColumn(name = "Photo", exportStringAsPicture = true)
private String photo;

// Multiple images in one cell (List<String>)
@PxlColumn(name = "Gallery", exportStringAsPicture = true)
private List<String> gallery;
```

### Column: Cell Styler Export

Stylers are delegated in the order column → sheet → workbook → built-in default.  
Use a built-in styler or implement your own (for the cascade/built-in list, see [Cell Stylers](#cell-stylers)).

```java
// Specify built-in stylers on the column (data cell + required header cell)
@PxlColumn(name = "Amount",
        exportColumnDataCellStyler = PxlDataCommaSeparatedNumericStyler.class,
        exportColumnRequiredHeaderCellStyler = PxlHeaderHorizontalCenterTextStyler.class)
private BigDecimal amount;
```

```java
// Workbook-wide data cell style (option) — applied if not specified separately on the sheet/column
PxlExportWorkbookOption.builder()
                       .exportWorkbookDataCellStyler(PxlDataThinBorderStyler.class)
                       .build();
```

```java
// Build your own — implement PxlStyler: modify the CellStyle and return the Font to apply to the cell
import io.github.hclimkr.pxl.styler.PxlStyler;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class BoldCenterStyler implements PxlStyler {

    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        final Font font = workbook.createFont();
        font.setBold(true);
        return font;                       // the returned Font is applied to the cell
    }
}
// @PxlColumn(name = "X", exportColumnDataCellStyler = BoldCenterStyler.class)
```

### Column: Fixed Column Order Export

Field declaration order does not guarantee column order. If order matters, set `exportOrder`.

```java
@PxlColumn(name = "A", exportOrder = "01")
private String a;

@PxlColumn(name = "B", exportOrder = "02")
private String b;
```

> `exportOrder` is a lexicographic string comparison.
> For numeric order, pad with leading zeros like `"01"`, `"02"`, … (`"2"` vs `"10"` → `"10"` comes first).

### Sheet: Multi-sheet Export

```java
// Call order = sheet order (Engineering, Sales, Departments); each sheet may take a different row class
pxl.exportExcel()
   .sheet(Employee.class, engineering, "Engineering")
   .sheet(Employee.class, sales, "Sales")
   .sheet(Department.class, departments, "Departments")
   .toFile(new File("company.xlsx"));

// Sample export works the same way
pxl.exportSampleExcel()
   .sheet(Employee.class, "Employees")
   .sheet(Department.class, "Departments")
   .toFile(new File("template.xlsx"));
```

### Sheet: Multi-sheet Import

```java
import io.github.hclimkr.pxl.builder.PxlExcelImportBuilder;

// 1) Workbook form: binds multiple sheets at once, one per @PxlSheet field
//    (Company = a @PxlWorkbook class with two @PxlSheet fields, employees/departments)
Company company = pxl.importExcel()
                     .workbook(Company.class)
                     .fromFile(file);

// 2) Sheet form: call once per sheet on the same builder and run each one (the return type may differ per sheet)
PxlExcelImportBuilder builder = pxl.importExcel();
List<Employee> employees = builder.sheet(Employee.class, "Employees")
                                  .fromFile(file);
List<Department> departments = builder.sheet(Department.class, "Departments")
                                      .fromFile(file);
```

### Sheet: Split Sheets by Grouping on a Column Value (Export)

Given `@PxlSheet(exportGroupingFieldName = "fieldName")`, it is split into multiple sheets by that value.

```java
@PxlSheet(name = "Employees", exportGroupingFieldName = "department")
private List<Employee> employees;   // exported to a separate sheet per department value
```

### Sheet: Import When Data Does Not Start at Row 1

Specify the starting row/column of the header and data as 1-based.
For the workbook form, give it directly on `@PxlSheet`; for the sheet form, give it via `PxlImportSheetOption`.

```java
// Workbook class: specified on @PxlSheet (header at row 3, data from row 4)
@PxlSheet(name = "Employees", importHeaderRowIndex = 3, importFirstDataRowIndex = 4)
private List<Employee> employees;
```

```java
// Sheet form: the row/column positions are on PxlImportSheetOption, which is placed in the workbook option's importSheetOptions
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                                                       .importHeaderRowIndex(3)
                                                       .importFirstDataRowIndex(4)
                                                       .build();

List<Employee> rows = pxl.importExcel()
                         .override(PxlImportWorkbookOption.builder()
                                                          .importSheetOptions(Arrays.asList(sheetOption))
                                                          .build())
                         .sheet(Employee.class, "Employees")
                         .fromFile(file);
```

> The end position and column range (`importLastDataRowIndex`, `importFirstDataColumnIndex`, `importLastDataColumnIndex`) are given the same way. All 1-based.
> The Streaming Reader has no `getFirstRowNum()`, so specifying the header row is required.

### Sheet: CSV Encoding and Delimiter Import

```java
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

List<Employee> rows = pxl.importCsv()
                         .override(PxlImportWorkbookOption.builder()
                                                          .importCsvCharset("US-ASCII") // default is "UTF-8" (only specify legacy encodings)
                                                          .importCsvDelimiter('\t')     // char (e.g., TSV). default ','
                                                          .build())
                         .sheet(Employee.class)
                         .fromFile(file);
```

> The default CSV encoding is `UTF-8`.
> For CSV in a different encoding (e.g., `US-ASCII`·`MS949`·`EUC-KR`), specify it with `importCsvCharset(...)` as above. `importCsvDelimiter` is a `char`, so use single quotes.

A CSV workbook is read as one file per sheet, so its sheets need not share an encoding or a delimiter.  
Both attributes therefore exist on `@PxlSheet` as well, and a sheet that names neither inherits the workbook value.

```java
@PxlWorkbook(importCsvCharset = "MS949")            // applies to every sheet that does not say otherwise
public class CompanyWorkbook {

    @PxlSheet(name = "Legacy")                       // inherits MS949
    private List<CharsetRow> legacy;

    @PxlSheet(name = "Modern", importCsvCharset = "UTF-8")   // reads this one file as UTF-8
    private List<CharsetRow> modern;
}

CompanyWorkbook workbook = pxl.importCsv()
                              .workbook(CompanyWorkbook.class)
                              .fromFiles(Arrays.asList(legacyCsv, modernCsv));
```

Resolution runs top to bottom, and the first level that names a value wins:

| Level | "not specified" is |
|-------|--------------------|
| `PxlImportSheetOption.importCsvCharset` / `importCsvDelimiter` | `null`, and `""` / `'\0'` too |
| `@PxlSheet(importCsvCharset)` / `(importCsvDelimiter)`         | `""` / `'\0'`     |
| `PxlImportWorkbookOption.importCsvCharset` / `importCsvDelimiter` | `null`, and `""` / `'\0'` too |
| `@PxlWorkbook(importCsvCharset)` / `(importCsvDelimiter)`      | `""` / `'\0'`     |
| built-in default                                               | `"UTF-8"` / `','` |

> Because "not specified" is a sentinel rather than the effective default, a sheet can name `"UTF-8"` or `','` explicitly to return to the default against a workbook that names something else.  
> In the sheet form (`sheet(...)`) there is no field to carry `@PxlSheet`, so a wildcard `PxlImportSheetOption` is the only sheet-level route — see *Option Structure*.  
> The same holds one level up: the sheet form builds its workbook metadata without a workbook class, so `@PxlWorkbook` is not read there either. In that form both annotation levels are inert and the options are what take effect.

### Sheet: CSV Encoding, Delimiter and BOM Export

The writing side mirrors the reading side, with a byte order mark added — a CSV file is written one per sheet, so
all three belong to the file rather than to the schema.

```java
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;

pxl.exportCsv()
   .sheet(Employee.class, employees, "Employees")
   .override(PxlExportWorkbookOption.builder()
                                    .exportCsvCharset("MS949")   // default is "UTF-8"
                                    .exportCsvDelimiter('\t')    // char (e.g., TSV). default ','
                                    .exportCsvBom(true)          // Boolean here; PxlOptionalBoolean on the annotations
                                    .build())
   .toFile(csvFile);
```

```java
@PxlWorkbook(exportCsvCharset = "MS949", exportCsvBom = PxlOptionalBoolean.TRUE)
public class CompanyWorkbook {

    @PxlSheet(name = "Legacy")                                    // inherits MS949 and the mark
    private List<CharsetRow> legacy;

    @PxlSheet(name = "Modern",
              exportCsvCharset = "UTF-8",
              exportCsvBom = PxlOptionalBoolean.FALSE)            // writes this one file without a mark
    private List<CharsetRow> modern;
}
```

Resolution runs the same five levels as the import side:

| Level | "not specified" is |
|-------|--------------------|
| `PxlExportSheetOption.exportCsvCharset` / `exportCsvDelimiter` / `exportCsvBom` | `null`, and `""` / `'\0'` too |
| `@PxlSheet(exportCsvCharset)` / `(exportCsvDelimiter)` / `(exportCsvBom)`       | `""` / `'\0'` / `UNSPECIFIED` |
| `PxlExportWorkbookOption.exportCsvCharset` / `exportCsvDelimiter` / `exportCsvBom` | `null`, and `""` / `'\0'` too |
| `@PxlWorkbook(exportCsvCharset)` / `(exportCsvDelimiter)` / `(exportCsvBom)`    | `""` / `'\0'` / `UNSPECIFIED` |
| built-in default                                                                | `"UTF-8"` / `','` / `false`  |

> `exportCsvBom` is the one that cannot use a plain `boolean` on the annotations. A `boolean` has no value left over
> to mean "not specified", so its `false` would be indistinguishable from silence and a sheet could never turn off a
> mark its workbook asked for. `PxlOptionalBoolean` (`UNSPECIFIED` · `TRUE` · `FALSE`) is what supplies the missing
> third value; the option classes keep a boxed `Boolean`, whose `null` already says it.  
> The mark itself is written only for UTF-8, UTF-16LE and UTF-16BE — see [Limitation](#limitation).  
> As with the import side, the sheet form (`sheet(...)`) reads neither annotation level, so the options are the route there.

### Workbook: Encrypt with Password Export

```java
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;

pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .override(PxlExportWorkbookOption.builder()
                                    .exportPassword("secret")
                                    .build())
   .toFile(new File("secured.xlsx"));
```

### Workbook: Streaming Reader

```java
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

// Streaming requires the header row position to be specified precisely, so provide it with a sheet option
PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                                                       .importHeaderRowIndex(1)   // 1-based
                                                       .build();

List<Employee> rows = pxl.importExcel()
                         .override(PxlImportWorkbookOption.builder()
                                                          .importUsingStreamReader(true)   // XLSX only
                                                          .importSheetOptions(Arrays.asList(sheetOption))
                                                          .build())
                         .sheet(Employee.class, "Employees")
                         .fromFile(bigFile);
```

> The Streaming Reader is XLSX-only, cannot evaluate formula cells, and has no `getFirstRowNum()`, so the header row must be specified precisely.

---

## Limitation

| Format | Max sheets | Max rows   | Max columns | Notes                                                         |
|--------|------------|------------|-------------|--------------------------------------------------------------|
| XLSX   | 100        | 1,048,576  | 16,384      | Can be read with the Streaming Reader without memory issues (GC overhead) |
| XLS    | 100        | 65,536     | 256         | The Streaming Reader is not supported, but non-streaming also has no memory issues |
| CSV    | 100        | 100,000    | 16,384      | One file is one sheet, so "sheets" counts the files passed to `fromFiles(...)`/`fromStreams(...)` — a CSV export writes a single sheet, so only the row/column caps apply there.<br/>The column cap matches XLSX so that a row class exporting to XLSX stays readable from CSV; the row cap is lower because an imported CSV is read into memory whole before the cap is checked |

> These are the limits `PxlFileFormat` carries (`getMaxExportRows()` and its siblings), so an engine is bound by the
> limits of the format it writes — `XSSF` and `SXSSF` alike.  
> The row/column figures for XLSX·XLS are the format's own; the sheet count, and every CSV figure, are limits PXL
> imposes rather than the format. Exceeding any of them raises `PxlDataException`.

- Automatic column width (`autoSizeColumn`) and large data  
  If you do not specify `exportColumnWidth`, the default (auto) applies, so on export POI `autoSizeColumn` measures every row cell of that column with font metrics (O(row count) per column).  
  With N columns and M rows, an O(N×M) measurement cost is added, which can become the dominant factor of performance degradation on large-data export.  
  If there are many rows, specifying a fixed width via `@PxlColumn(exportColumnWidth = ...)` or an option is recommended.

- Export heap and `SXSSF`  
  `XSSF` (the default) builds the whole workbook as an object graph before a single byte is written, so the heap it holds is far larger than the file it finally produces.  
  `SXSSF` writes the same `.xlsx` while keeping only a sliding window of rows in memory (`exportSXSSFRowAccessWindowSize`, default 100) and spilling the rest to temp files, which makes the heap roughly independent of the row count.  
  It applies to XLSX only — it has no effect on the `HSSF` engine (`.xls`).  
  Note that a column left at automatic width must be tracked in order to be measured, and a tracked column stays in memory, which eats into the saving. Pair `SXSSF` with a fixed `exportColumnWidth`.

- CSV export builds its output before writing, spilling to a temporary file when it is large  
  A CSV export renders the whole file before the destination is opened, which is what makes a failure leave no file behind and lets `toStream(...)` write in one pass. Up to `PxlConstants.EXPORT_MEMORY_THRESHOLD_OF_CSV` (4 MiB) that output is held in memory; past it the rest continues into a temporary file, so the heap a CSV export needs does not grow with the output. The file is created under `java.io.tmpdir` with the prefix `pxl-csv-export-`, and is deleted before the call returns whether it succeeded or failed.  
  Three consequences of that spill are worth knowing. A large export needs free disk space, and running out of it fails as a `PxlIOException`. Its contents are written unencrypted, which matters because a CSV export refuses `exportPassword` rather than encrypting — if the rows are sensitive, point `java.io.tmpdir` somewhere you trust. And a JVM killed mid-export leaves the file behind, since only a normal return or a thrown exception reaches the cleanup.  
  Should memory run out anyway, the `OutOfMemoryError` reaches you as itself rather than as a `PxlException`: an `Error` is not an `Exception`, so the terminal method's normalization does not cover it — and catching it would be wrong, since wrapping allocates in the very condition that ran out of memory. The rows you pass in are still held in full, so that remains the term to size the heap for.

- Settings a CSV export ignores  
  Everything that only describes how a cell looks or what a workbook contains: `@PxlColumn(exportColumn*Styler, exportColumnWidth, exportOptionItems, exportEnumDropDownListStyle)`, `@PxlSheet(exportSheet*Styler, exportRowHeightInPoints, exportColumnFilter, exportGroupingFieldName)`, `@PxlWorkbook(exportWorkbook*Styler, exportExcelEngine, exportSXSSFRowAccessWindowSize)`, and the header freeze pane. Grouping being ignored means the rows are written in the order given, not gathered per group.  
  Two attributes still put their value in the file rather than dropping it: `exportStringAsFormula` writes the text as it stands, leading `=` and all (nothing is evaluated), and `exportStringAsPicture` writes the image location instead of embedding a picture — so a path or URL you expected to disappear is disclosed.  

- CSV byte order mark  
  `exportCsvBom` is honored for UTF-8, UTF-16LE and UTF-16BE only. Any other charset writes no mark even when it is asked for, with neither an exception nor a warning: `UTF-16` has its encoder emit one already, and a non-Unicode charset such as EUC-KR cannot encode U+FEFF and would replace it with `?`, corrupting the first header field. Since there is no way to notice this at runtime, check the charset here before concluding a mark went missing.

---

## Common Pitfalls Checklist

- ✅ **A no-arg constructor is required on the DTO.**  
  Row objects are created by reflection on import.  
  When using Lombok, putting only `@AllArgsConstructor` removes the no-arg constructor, so add `@NoArgsConstructor` as well.
- ✅ **Name matching**  
  The `name` of `@PxlColumn`/`@PxlSheet` (the field name if unspecified) must match the actual header/sheet name. Whitespace is ignored on both. A sheet name ignores case as well while a column header is matched case-sensitively.
- ✅ **Columns whose names do not match**  
  If required (`@NotNull`/`@NotEmpty`/`@NotBlank`), an exception is raised; if not required, it is silently excluded.
- ✅ **Indices are 1-based**  
  `importHeaderRowIndex`, `exportFirstDataColumnIndex`, etc., are all 1-based.  
  The value received by `@PxlRowIndex` is likewise 1-based — the spreadsheet row number of the imported row.
- ✅ **The default CSV encoding is `UTF-8`**  
  Specify other encodings (`US-ASCII`·`MS949`·`EUC-KR`, etc.) with `importCsvCharset(...)`, or `exportCsvCharset(...)` when writing.
  In the workbook form each sheet is its own file, so a sheet that differs from the rest can name its own `@PxlSheet(importCsvCharset)` / `(importCsvDelimiter)` — and, on the writing side, `(exportCsvCharset)` / `(exportCsvDelimiter)` / `(exportCsvBom)` — instead of forcing the whole workbook onto one setting.
- ✅ **A CSV export quotes only what needs quoting**  
  A value such as `"010"` or `"1E+10"` is written unquoted, so Excel reads it back as the number `10` or `1e10`. PXL reads it back as the string it was; the loss happens in the spreadsheet application, not in the file. There is no quoting-policy setting yet — write such values with a `pattern` if the display matters.
- ✅ **`long` / `BigInteger` / `BigDecimal` precision**  
  Large numbers may lose precision in a numeric cell (double, 2^53 limit). To preserve them exactly, output as a string cell with a `pattern` or use `BigInteger`/`BigDecimal`.
- ✅ **Reuse `Pxl`.**  
  `new Pxl()` has a validation-bootstrap cost, so keep it as a singleton/Spring bean (thread-safe).
- ✅ **Streaming Reader constraints**  
  XLSX-only, cannot evaluate formulas, must specify the header row position precisely.
- ✅ **Large-export performance**  
  When the column width is unspecified, `autoSizeColumn` measures every row. If there are many rows, give a fixed width via `@PxlColumn(exportColumnWidth = ...)`.

---

## License

This project is distributed under the **[Apache License 2.0](../LICENSE)** (for the full text, see [License in README](../README.md#license)).
