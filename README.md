**English** · [한국어](README_ko.md)

PXL - Java Excel & CSV to POJO Mapping Library
=============================

[![Build](https://github.com/hclimkr/pxl/actions/workflows/build.yml/badge.svg?branch=main)](https://github.com/hclimkr/pxl/actions/workflows/build.yml)
[![Maven Central](https://img.shields.io/maven-central/v/io.github.hclimkr/pxl-javax?label=Maven%20Central)](https://central.sonatype.com/artifact/io.github.hclimkr/pxl-javax)
[![Javadoc](https://javadoc.io/badge2/io.github.hclimkr/pxl-javax/javadoc.svg)](https://javadoc.io/doc/io.github.hclimkr/pxl-javax)
[![Java](https://img.shields.io/badge/Java-8%2B%20%2F%2017%2B-orange.svg)](#setup)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)

PXL is an **annotation-driven, bidirectional binding between spreadsheets and Java objects**,
built on top of Apache POI and Apache Commons CSV, and supports Java 8 and above.

Read an uploaded `.xlsx`, `.xls` or `.csv` file straight into a `List<Employee>`, and write a
`List<Employee>` straight back out as an Excel or CSV download — no `Row`/`Cell` loops, no manual
type conversion, no separate reader and writer to keep in sync. You annotate the DTO once and that
one declaration drives both directions.

- Import: XLSX · XLS · CSV → Java objects
- Export: Java objects → XLSX · XLS · streaming XLSX · CSV
- Only fields/classes marked with the dedicated annotations are bound.

```java
List<Employee> employees = pxl.importExcel()
                              .sheet(Employee.class, "Employees")
                              .fromFile(new File("employees.xlsx"));
```

For details such as supported variable types, the full set of options, and constraints, refer to [docs/reference.md](docs/reference.md).

> [!WARNING]
> **Pre-1.0: the public API is still moving.** Under Semantic Versioning a `0.y.z` release makes no compatibility
> promise, and PXL uses that room: a minor release may rename, move or remove a public type or method without a
> deprecation cycle — several already have. Pin an exact version rather than a range, and read the
> [CHANGELOG](CHANGELOG.md) before upgrading; every such change is listed there and marked breaking.

## Table of Contents

1. [Features](#features)
2. [Setup](#setup)
3. [Defining DTO Classes](#defining-dto-classes)
4. [Usage at a Glance](#usage-at-a-glance)
5. [Export (Objects → Excel)](#export-objects--excel)
6. [Export Sample (Class → Sample Excel)](#export-sample-class--sample-excel)
7. [Export (Objects → CSV)](#export-objects--csv)
8. [Export Sample (Class → Sample CSV)](#export-sample-class--sample-csv)
9. [Import (Excel → Objects)](#import-excel--objects)
10. [Import (CSV → Objects)](#import-csv--objects)
11. [FAQ](#faq)
12. [Build & Contributing](#build--contributing)
13. [License](#license)

---

## Features

- **POJO to Excel and Excel to POJO from one declaration** — the same `@PxlColumn` mapping is read by
  the exporter and the importer, so a round trip cannot drift apart.
- **XLSX, XLS and streaming XLSX** — pick the POI engine (`XSSF` / `HSSF` / `SXSSF`) with a single
  annotation attribute; the file format, extension and sheet/row/column limits follow from it.
- **CSV in the same model** — the same annotations, the same converters and the same column order
  write a CSV instead of a workbook, and read it straight back.
- **Declarative mapping only** — `@PxlColumn` on a field, `@PxlSheet` on a collection field,
  `@PxlWorkbook` on the class. Nothing else in the DTO is touched. Header names match by name, so
  column order in the file is free and unknown columns are ignored.
- **Around 30 field types out of the box** — primitives and their wrappers, `String`, `BigDecimal`,
  `BigInteger`, `enum`, `LocalDate` / `LocalDateTime` / `LocalTime` / `Date`, `Duration` / `Period`,
  `UUID`, `Collection`, and anything else through `@PxlExportConverter` / `@PxlImportConverter`.
- **Bean Validation per row** — standard `javax.validation` / `jakarta.validation` constraints
  declared on the DTO are enforced while binding, alongside PXL's own `@PxlByteSize`.
- **Excel-side presentation** — cell stylers, column widths, row heights, freeze panes, auto filters,
  dropdown validation, row grouping, embedded pictures and password-protected workbooks.
- **Built for large files** — export can run on SXSSF and a CSV export spills past 4 MiB to a
  temporary file instead of growing the heap, while import can run on `excel-streaming-reader`.
- **Sample templates from a class alone** — generate a header row plus one filled example row as an
  `.xlsx` or `.csv` form to hand out and collect back.
- **Localization on two independent channels** — translate sheet and column names for the content,
  and switch the language of PXL's own diagnostics and exception messages (English / Korean bundled).
- **Two variants, one source** — `pxl-javax` for Java 8+ and `pxl-jakarta` for Java 17+.

---

## Setup

Add only the variant that matches your environment to your dependencies.
- `pxl-javax` (Java 8+, `javax.*`)
- `pxl-jakarta` (Java 17+, `jakarta.*`)

### Maven

```xml
<!-- javax variant (Java 8+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-javax</artifactId>
    <version>0.9.6</version>
</dependency>
```

```xml
<!-- jakarta variant (Java 17+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-jakarta</artifactId>
    <version>0.9.6</version>
</dependency>
```

### Gradle

```groovy
// javax variant (Java 8+)
implementation 'io.github.hclimkr:pxl-javax:0.9.6'
```

```groovy
// jakarta variant (Java 17+)
implementation 'io.github.hclimkr:pxl-jakarta:0.9.6'
```

---

## Defining DTO Classes

### Row Class

A row class maps each field to a header with `@PxlColumn`.

```java
import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.time.LocalDate;

@Getter                     // (optional) for your convenience — PXL binds fields directly, so a getter is not required.
@Setter                     // (optional) for your convenience — PXL binds fields directly, so a setter is not required.
@NoArgsConstructor          // (required) no-arg constructor is required for import
public class Employee {

    @PxlRowIndex            // (optional) 1-based spreadsheet row number. Type: byte/short/int/long + wrapper classes (Byte/Short/Integer/Long)
    private Integer rowIndex;

    @PxlColumn(name = "Name", exportSample = "John Doe")
    private String name;

    @PxlColumn(name = "Age", exportSample = "25")
    private Integer age;

    @PxlColumn(name = "Salary", exportSample = "45000")
    private Long salary;

    @PxlColumn(name = "Active", exportSample = "true")
    private Boolean active;

    @PxlColumn(name = "HireDate", pattern = "yyyy-MM-dd", exportSample = "2024-03-01")
    private LocalDate hireDate;

    @PxlColumn(name = "Grade", exportSample = "C")
    private Grade grade;
}
```

- If `name` is omitted, the field name becomes the column name.
- `name` must match the actual header for binding to occur (whitespace is ignored, case is significant).
- `exportSample` is the example value that goes into [Export Sample](#export-sample-class--sample-excel) (it has no effect on a regular export).
- `exportSample` is written as a `String` but parsed into the column type, so it must be a value that type accepts (`PxlCellCodecException` otherwise). A column left without one gets `exportNullString` (default `""`) in the sample row.

`Grade` is a user-defined enum used in the examples.

```java
public enum Grade {
    A, B, C, F
}
```

The row class used as the second sheet in the multi-sheet example is defined the same way.

```java
import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter                     // (optional) for your convenience — PXL binds fields directly, so a getter is not required.
@Setter                     // (optional) for your convenience — PXL binds fields directly, so a setter is not required.
@NoArgsConstructor          // (required) no-arg constructor is required for import
public class Department {

    @PxlColumn(name = "Code")
    private String code;

    @PxlColumn(name = "DepartmentName")
    private String departmentName;

    @PxlColumn(name = "Headcount")
    private int headcount;
}
```

### Workbook Class (Multiple Sheets in One Object)

Each sheet field is a `Collection` type and is bound with `@PxlSheet`.

```java
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter                     // (optional) for your convenience — PXL binds fields directly, so a getter is not required.
@Setter                     // (optional) for your convenience — PXL binds fields directly, so a setter is not required.
@NoArgsConstructor          // (required) no-arg constructor is required for import
public class Company {

    @PxlWorkbookName        // (optional) a String field to hold the workbook name
    private String workbookName;

    @PxlSheet(name = "Employees")
    private List<Employee> employees;

    @PxlSheet(name = "Departments")
    private List<Department> departments;
}
```

---

## Usage at a Glance

### Export

```java
import io.github.hclimkr.pxl.Pxl;

import java.io.File;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

// Create Pxl once and reuse it (thread-safe, stateless — singleton/Spring bean recommended)
Pxl pxl = new Pxl();

// Prepare an Employee row object
Employee alice = new Employee();
alice.setName("Alice");
alice.setAge(30);
alice.setSalary(50_000L);
alice.setActive(true);
alice.setHireDate(LocalDate.of(2020, 1, 15));
alice.setGrade(Grade.A);

// Prepare an Employee row object
Employee bob = new Employee();
bob.setName("Bob");
bob.setAge(42);
bob.setSalary(72_000L);
bob.setActive(false);
bob.setHireDate(LocalDate.of(2018, 6, 1));
bob.setGrade(Grade.B);

// Prepare the Employees sheet object
List<Employee> employees = Arrays.asList(alice, bob);

// Prepare a Department row object
Department eng = new Department();
eng.setCode("ENG");
eng.setDepartmentName("Engineering");
eng.setHeadcount(12);

// Prepare a Department row object
Department sal = new Department();
sal.setCode("SAL");
sal.setDepartmentName("Sales");
sal.setHeadcount(8);

// Prepare the Departments sheet object
List<Department> departments = Arrays.asList(eng, sal);

// Prepare the Company workbook object
Company company = new Company();
company.setEmployees(employees);
company.setDepartments(departments);

// Export: Company workbook object → Excel file
pxl.exportExcel()
   .workbook(company)
   .toFile(new File("company.xlsx"));
```

### Import

```java
import io.github.hclimkr.pxl.Pxl;

import java.io.File;

// Create Pxl once and reuse it (thread-safe, stateless — singleton/Spring bean recommended)
Pxl pxl = new Pxl();

// Import: Excel file → Company workbook object
Company company = pxl.importExcel()
                     .workbook(Company.class)
                     .fromFile(new File("company.xlsx"));
```
Every operation is handled through a single method chain like the examples above. The first method name indicates the direction of the operation (export/import) and the format (Excel/CSV), then you specify the target, and it is executed in the final (execute) method.

| Use case          | Method chain (start → configure → execute)                                                                                                                                                      |
|-------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Excel export      | `pxl.exportExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                                     |
| Sample Excel export | `pxl.exportSampleExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                               |
| CSV export        | `pxl.exportCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                          |
| Sample CSV export | `pxl.exportSampleCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                    |
| Excel import      | `pxl.importExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromStream(InputStream)`                                                                                    |
| CSV import        | `pxl.importCsv()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromFiles(List<File>)` / `.fromStream(String, InputStream)` / `.fromStreams(List<String>, List<InputStream>)` |

---

## Export (Objects → Excel)

Configure the content with `workbook(...)` or `sheet(...)` (specify only one of the two forms; specifying both throws an exception),
and finally output with one of `toFile(File)` / `toStream(OutputStream)` / `toWorkbook()`.

### Workbook Object → Excel

```java
Company company = ...;

pxl.exportExcel()
   .workbook(company)
   .toFile(new File("company.xlsx"));
```

### Single Sheet Object → Excel

```java
pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .toFile(new File("employees.xlsx"));
```

### Export Result Excel Layout (Single Sheet)

Exporting the `Employee` list above with `sheet(Employee.class, employees, "Employees")` produces an "Employees" sheet like the following (column letters A–F and row numbers are as shown on the Excel screen).

|       | A     | B   | C      | D      | E          | F     |
|-------|-------|-----|--------|--------|------------|-------|
| **1** | Name  | Age | Salary | Active | HireDate   | Grade |
| **2** | Alice | 30  | 50000  | true   | 2020-01-15 | A     |
| **3** | Bob   | 42  | 72000  | false  | 2018-06-01 | B     |

- Row 1 is the header row, and rows from 2 onward are data rows.
- Cell types: `Name` is text, `Age` and `Salary` are numeric, `Active` is `true`/`false` text, `HireDate` and `Grade` are text.
- `@PxlRowIndex` (`rowIndex`) is not a column, so it is not written to the sheet.

### Multiple Sheet Object Form

```java
pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .sheet(Department.class, departments, "Departments")
   .toFile(new File("company.xlsx"));
```

### Export Result Excel Layout (Multiple Sheets)

Exporting a workbook object (`Company`) or repeated `.sheet(...)` calls produces multiple sheets in a single `.xlsx` — one `@PxlSheet` field (or each `.sheet(...)`) is one sheet. The example above creates two sheet tabs `[ Employees | Departments ]`.

- The "Employees" sheet — identical to the `Employee` table in [Export Result Excel Layout (Single Sheet)](#export-result-excel-layout-single-sheet) above.
- The "Departments" sheet is produced like the following.

|       | A    | B              | C         |
|-------|------|----------------|-----------|
| **1** | Code | DepartmentName | Headcount |
| **2** | ENG  | Engineering    | 12        |
| **3** | SAL  | Sales          | 8         |

- The workbook class's `@PxlWorkbookName` `String` field is used for the workbook name rather than a column (cell), so it does not appear in any sheet.
- Import is symmetric — reading an `.xlsx` with these two sheets via `workbook(Company.class)` matches the sheet names and fills each `@PxlSheet` field (`employees`, `departments`).

### Output Targets

```java
// 1) File (PXL opens and closes the file internally)
pxl.exportExcel()
   .sheet(Row.class, rows, "S")
   .toFile(new File("out.xlsx"));
```

```java
// 2) Stream (e.g. an HTTP response) — PXL does not close the stream; the caller closes it
try (OutputStream os = response.getOutputStream()) {
    pxl.exportExcel()
       .sheet(Row.class, rows, "S")
       .toStream(os);
}
```

```java
// 3) POI Workbook — the caller closes the returned workbook
Workbook workbook = pxl.exportExcel()
                       .sheet(Row.class, rows, "S")
                       .toWorkbook();
```

---

## Export Sample (Class → Sample Excel)

From a class alone, this creates a sample template with a header row plus one data row filled with each column's example value (`exportSample`).
The final output method (`toFile`/`toStream`/`toWorkbook`) is the same as a regular export.

```java
// Workbook class form
pxl.exportSampleExcel()
   .workbook(Company.class)
   .toFile(new File("sample.xlsx"));
```

```java
// Single sheet form
pxl.exportSampleExcel()
   .sheet(Employee.class, "Employees")
   .toFile(new File("sample.xlsx"));
```

```java
// Multiple sheet form
pxl.exportSampleExcel()
   .sheet(Employee.class, "Employees")
   .sheet(Department.class, "Departments")
   .toFile(new File("sample.xlsx"));
```

### Export Sample Result Excel Layout

Creating `exportSampleExcel().sheet(Employee.class, "Employees")` with the `Employee` above produces an "Employees" sheet with a header row plus one sample data row.

|       | A        | B   | C      | D      | E          | F     |
|-------|----------|-----|--------|--------|------------|-------|
| **1** | Name     | Age | Salary | Active | HireDate   | Grade |
| **2** | John Doe | 25  | 45000  | true   | 2024-03-01 | C     |

- For multiple sheets (`Company` or `.sheet(...).sheet(...)`), each sheet gets a header row plus one sample data row in the same manner.

---

## Export (Objects → CSV)

A CSV file holds one sheet, so this builder has the sheet form only — there is no `workbook(...)` — and the final
method writes that one sheet. Everything about the values themselves is shared with the Excel export: the same
annotations, the same converters, the same column order.

```java
pxl.exportCsv()
   .sheet(Employee.class, employees, "Employees")
   .toFile(new File("employees.csv"));
```

```csv
Name,Age,Salary,Active,HireDate,Grade
Alice,30,50000,true,2020-01-15,A
Bob,42,72000,false,2018-06-01,B
```

### Encoding, Delimiter and Byte Order Mark

A CSV workbook is one file per sheet, so these belong to the file rather than to the schema and are settled at the
sheet level. The sheet form binds no `@PxlSheet` field, so a wildcard `PxlExportSheetOption` is the sheet-level
route; a workbook-level option covers every sheet at once.

```java
PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
        .exportCsvCharset("EUC-KR")
        .exportCsvDelimiter('\t')
        .exportCsvBom(true)
        .build();

pxl.exportCsv()
   .sheet(Employee.class, employees, "Employees")
   .override(option)
   .toFile(new File("employees.csv"));
```

- A byte order mark is written only for UTF-8, UTF-16LE and UTF-16BE. Any other charset leaves it out silently —
  UTF-16 has its encoder write one already, and a non-Unicode charset such as EUC-KR cannot encode it at all and
  would corrupt the first field.
- The header line is always written; there is no switch to turn it off.

### What CSV Does Not Carry

Settings that describe how a cell looks or what a workbook contains are ignored: stylers, column widths, row
heights, freeze panes, auto-filters, dropdowns, the engine and its streaming window, and `exportGroupingFieldName`
(the rows stay in the order given). Two of them still put their value in the file — `exportStringAsFormula` writes
the text as it stands, leading `=` and all, and `exportStringAsPicture` writes the image location instead of
embedding a picture.

`exportPassword` is the one that is refused rather than ignored: CSV cannot be encrypted, and writing plaintext
instead would be a leak.

---

## Export Sample (Class → Sample CSV)

From a class alone, this creates a sample template with a header line plus one data line filled with each column's
example value (`exportSample`) — the CSV counterpart of
[Export Sample (Class → Sample Excel)](#export-sample-class--sample-excel).
A CSV file holds one sheet, so this builder too has the sheet form only, and the final method is
`toFile`/`toStream` (there is no `toWorkbook`).

```java
pxl.exportSampleCsv()
   .sheet(Employee.class, "Employees")
   .toFile(new File("employees-template.csv"));
```

### Export Sample Result CSV Layout

Creating `exportSampleCsv().sheet(Employee.class, "Employees")` with the `Employee` above produces the following file.

```csv
Name,Age,Salary,Active,HireDate,Grade
John Doe,25,45000,true,2024-03-01,C
```

- A column without `exportSample` gets `exportNullString` (default `""`), so its field is left empty.
- `importCsv()` reads this file straight back, which makes it a form to hand out and collect filled in.
- A CSV file carries no sheet name of its own, so the name given here only picks the matching sheet-level option
  and labels error messages.
- Encoding, delimiter and byte order mark are settled exactly as in
  [Export (Objects → CSV)](#export-objects--csv), and the same settings are ignored or refused there.

---

## Import (Excel → Objects)

Start with `importExcel()`, configure the target to read as a workbook form (`workbook(Class)`) or a sheet form (`sheet(Class, candidateSheetNames...)`),
and specify the final (execute) method (`fromFile`/`fromStream`) to parse it on the spot.

### Import Source Excel Layout

If the "Employees" sheet to read has the layout below, it is bound to `Employee` (the same structure as [Export Result Excel Layout](#export-result-excel-layout-single-sheet) above — round trip).

|       | A     | B   | C      | D      | E          | F     |
|-------|-------|-----|--------|--------|------------|-------|
| **1** | Name  | Age | Salary | Active | HireDate   | Grade |
| **2** | Alice | 30  | 50000  | true   | 2020-01-15 | A     |
| **3** | Bob   | 42  | 72000  | false  | 2018-06-01 | B     |

- Matching occurs when the header name matches `@PxlColumn(name=...)` (whitespace is ignored, case is significant). Column order is free, and columns not in the definition are ignored.
- `@PxlRowIndex` (`rowIndex`) is filled automatically with the 1-based spreadsheet row number of each read row.

### Workbook Form (@PxlWorkbook Object)

```java
Company company = pxl.importExcel()
                     .workbookName("Acme")            // (optional) value to fill into the @PxlWorkbookName field
                     .workbook(Company.class)
                     .fromFile(new File("company.xlsx"));
```

### Sheet Form (Returns Collection)

```java
// Return as a List from a file
List<Employee> rows = pxl.importExcel()
                         .sheet(Employee.class, "Employees")
                         .fromFile(new File("employees.xlsx"));
```

```java
// Return as a Set from a file
Set<Employee> set = pxl.importExcel()
                       .sheet(Employee.class, Set.class, "Employees")
                       .fromFile(new File("employees.xlsx"));
```

```java
// Return as a List from a stream (PXL does not close the stream; the caller closes it)
try (InputStream is = new FileInputStream("employees.xlsx")) {
    List<Employee> rows2 = pxl.importExcel()
                              .sheet(Employee.class, "Employees")
                              .fromStream(is);
}
```

---

## Import (CSV → Objects)

This works the same way as Excel import (start → configure workbook/sheet → call the final (execute) method). For CSV, the file name (without extension) becomes the sheet name, and the final (execute) methods expand to `fromFile`/`fromFiles`/`fromStream`/`fromStreams`.
The sheet form uses `sheet(Class)` without candidate sheet name arguments (single CSV), while the workbook form groups multiple CSVs by sheet.

### Import Source CSV File Layout

CSV is plain text with the first line as the header and subsequent lines as data (unlike Excel, there are no cell types, so dates and booleans are also written as strings).
The name of the file with its extension removed becomes the sheet name. In the workbook form it has to match the `@PxlSheet` name to bind, with whitespace and case ignored.

```text
Name,Age,Salary,Active,HireDate,Grade
Alice,30,50000,true,2020-01-15,A
Bob,42,72000,false,2018-06-01,B
```

Reading the `Employees.csv` above with `sheet(Employee.class).fromFile(...)` produces the `List<Employee>` below.

| Name  | Age | Salary | Active | HireDate   | Grade |
|-------|-----|--------|--------|------------|-------|
| Alice | 30  | 50000  | true   | 2020-01-15 | A     |
| Bob   | 42  | 72000  | false  | 2018-06-01 | B     |

- The default encoding is `UTF-8` and the delimiter is `,`.
- Header name matching and value interpretation (boolean tokens, free column order, ignoring columns not in the definition, etc.) are the same as Excel import (see [Import Source Excel Layout](#import-source-excel-layout) above).
- When grouping multiple CSVs in the workbook form, each file becomes one sheet — e.g. `Departments.csv` → sheet `Departments`.

### Workbook Form (Grouping Multiple CSVs into Sheets)

```java
// Multiple files → each file is one sheet
Company company = pxl.importCsv()
                     .workbookName("Acme")
                     .workbook(Company.class)
                     .fromFiles(Arrays.asList(employeesCsv, departmentsCsv));
```

```java
// Multiple streams (list of names + list of streams)
Company company = pxl.importCsv()
                     .workbook(Company.class)
                     .fromStreams(names, streams);
```

### Sheet Form (Single CSV)

```java
// File (no sheet name argument — CSV is a single table)
List<Employee> employees = pxl.importCsv()
                              .sheet(Employee.class)
                              .fromFile(new File("Employees.csv"));
```

```java
// Stream (name + stream)
List<Employee> employees = pxl.importCsv()
                              .sheet(Employee.class)
                              .fromStream("Employees", inputStream);
```

---

For more details (per-type behavior, the full set of annotation attributes, i18n, stylers, exceptions, etc.), refer to [docs/reference.md](docs/reference.md).

---

## FAQ

**How do I read an Excel file into a list of Java objects?**
Annotate the DTO fields with `@PxlColumn`, then call
`pxl.importExcel().sheet(Employee.class, "Employees").fromFile(file)`. It returns a `List<Employee>`
with every cell already converted to the field's type — see [Import (Excel → Objects)](#import-excel--objects).

**How do I return an Excel download from a Spring controller?**
Write to the servlet output stream with `toStream(...)`. PXL does not close the stream, so the caller
keeps control of the response — see [Output Targets](#output-targets). The matching content type and
file extension are available from `PxlFileFormat`.

**Can it handle files too large to fit in memory?**
Export can use the `SXSSF` engine, and a CSV export moves to a temporary file once it passes 4 MiB.
Import can run on `excel-streaming-reader` with `@PxlWorkbook(importUsingStreamReader = true)` (XLSX
only, and formula cells cannot be evaluated).

**Does it read `.xls` as well as `.xlsx`?**
Yes. Export chooses it with `@PxlWorkbook(exportExcelEngine = PxlExcelEngine.HSSF)` for `.xls`, and
import detects the format from the file itself.

**Does it handle CSV too?**
Yes, through `exportCsv()` / `importCsv()`, with the same annotations as Excel. Character set,
delimiter and byte order mark are configurable, and multiple CSV files can be read as one workbook
with one sheet per file.

**What if a column's type is not one PXL knows?**
Write a `@PxlExportConverter` / `@PxlImportConverter` method pair on the class and PXL will route that
column through it.

**Which artifact do I need, `pxl-javax` or `pxl-jakarta`?**
`pxl-javax` for `javax.validation` on Java 8 or later, `pxl-jakarta` for `jakarta.validation` on Java
17 or later. They are the same library — add exactly one.

**Is Bean Validation required?**
No. If no implementation is on the classpath, PXL logs a warning and simply skips validation instead
of failing.

**Is a `Pxl` instance safe to share?**
Yes — it is stateless and thread-safe, so create it once and reuse it as a singleton or a Spring bean.

---

## Build & Contributing

The source lives only in `pxl-javax`, and `pxl-jakarta` is generated by string substitution at build time.  
This repository accepts issue reports and suggestions only — see [CONTRIBUTING.md](CONTRIBUTING.md).

---

## License

This project is distributed under the [Apache License 2.0](LICENSE).

```
Copyright 2026 hclim

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
