**English** · [한국어](README_ko.md)

PXL
=============================

[![Build](https://github.com/hclimkr/pxl/actions/workflows/build.yml/badge.svg?branch=main)](https://github.com/hclimkr/pxl/actions/workflows/build.yml)
[![Maven Central](https://img.shields.io/maven-central/v/io.github.hclimkr/pxl-javax?label=Maven%20Central)](https://central.sonatype.com/artifact/io.github.hclimkr/pxl-javax)
[![Javadoc](https://javadoc.io/badge2/io.github.hclimkr/pxl-javax/javadoc.svg)](https://javadoc.io/doc/io.github.hclimkr/pxl-javax)
[![Java](https://img.shields.io/badge/Java-8%2B%20%2F%2017%2B-orange.svg)](#setup)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)

PXL is an **annotation-driven, bidirectional binding between spreadsheets and Java objects**,
built on top of Apache POI and Apache Commons CSV, and supports Java 8 and above.

- Import: XLSX · XLS · CSV → Java objects
- Export: Java objects → XLSX · XLS · streaming XLSX
- Only fields/classes marked with the dedicated annotations are bound.

For details such as supported variable types, the full set of options, and constraints, refer to [docs/reference.md](docs/reference.md).

## Table of Contents

1. [Setup](#setup)
2. [Defining DTO Classes](#defining-dto-classes)
3. [Usage at a Glance](#usage-at-a-glance)
4. [Export (Objects → Excel)](#export-objects--excel)
5. [Export Sample (Class → Sample Excel)](#export-sample-class--sample-excel)
6. [Import (Excel → Objects)](#import-excel--objects)
7. [Import (CSV → Objects)](#import-csv--objects)
8. [Build & Contributing](#build--contributing)
9. [License](#license)

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
    <version>0.9.2</version>
</dependency>
```

```xml
<!-- jakarta variant (Java 17+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-jakarta</artifactId>
    <version>0.9.2</version>
</dependency>
```

### Gradle

```groovy
// javax variant (Java 8+)
implementation 'io.github.hclimkr:pxl-javax:0.9.2'
```

```groovy
// jakarta variant (Java 17+)
implementation 'io.github.hclimkr:pxl-jakarta:0.9.2'
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
