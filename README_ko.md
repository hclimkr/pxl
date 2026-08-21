[English](README.md) · **한국어**

PXL - 자바 엑셀 · CSV 객체 매핑 라이브러리
=============================

[![Build](https://github.com/hclimkr/pxl/actions/workflows/build.yml/badge.svg?branch=main)](https://github.com/hclimkr/pxl/actions/workflows/build.yml)
[![Maven Central](https://img.shields.io/maven-central/v/io.github.hclimkr/pxl-javax?label=Maven%20Central)](https://central.sonatype.com/artifact/io.github.hclimkr/pxl-javax)
[![Javadoc](https://javadoc.io/badge2/io.github.hclimkr/pxl-javax/javadoc.svg)](https://javadoc.io/doc/io.github.hclimkr/pxl-javax)
[![Java](https://img.shields.io/badge/Java-8%2B%20%2F%2017%2B-orange.svg)](#구성)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)

PXL은 **애노테이션 기반으로 스프레드시트와 자바 객체를 양방향 바인딩하는 라이브러리**다.
Apache POI와 Apache Commons CSV 위에 구축되었으며, Java 8 이상을 지원한다.

업로드된 `.xlsx` · `.xls` · `.csv` 파일을 곧바로 `List<Employee>`로 읽고, `List<Employee>`를 곧바로
엑셀 · CSV 다운로드로 내보낸다 — `Row`/`Cell` 순회도, 수동 타입 변환도, 따로 관리해야 하는 리더와
라이터도 없다. DTO에 애노테이션을 한 번 붙이면 그 선언 하나가 양방향을 모두 구동한다.

- Import: XLSX · XLS · CSV → 자바 객체
- Export: 자바 객체 → XLSX · XLS · 스트리밍 XLSX · CSV
- 전용 애노테이션이 붙은 필드/클래스만 바인딩된다.

```java
List<Employee> employees = pxl.importExcel()
                              .sheet(Employee.class, "Employees")
                              .fromFile(new File("employees.xlsx"));
```

지원 변수 타입 · 전체 옵션 · 제약 등 상세 내용은 [docs/reference_ko.md](docs/reference_ko.md)를 참고한다.

> [!WARNING]
> **1.0 이전이라 공개 API가 아직 확정되지 않았다.** 유의적 버전(SemVer)에서 `0.y.z`는 호환성을 약속하지 않으며,
> PXL도 그 여지를 쓰고 있다 — 마이너 릴리스에서도 공개 타입·메서드가 deprecation 단계 없이 이름이 바뀌거나
> 옮겨지거나 사라질 수 있고, 실제로 그런 변경이 여러 번 있었다. 버전 범위 대신 정확한 버전을 고정하고,
> 올리기 전에 [CHANGELOG](CHANGELOG.md)를 확인한다. 해당 변경은 모두 거기에 breaking으로 표시되어 있다.

## 목차

1. [주요 기능](#주요-기능)
2. [구성](#구성)
3. [객체 DTO 정의](#객체-dto-정의)
4. [한눈에 보는 사용법](#한눈에-보는-사용법)
5. [Export (객체 → 엑셀)](#export-객체--엑셀)
6. [Export 샘플 (클래스 → 샘플 엑셀)](#export-샘플-클래스--샘플-엑셀)
7. [Export (객체 → CSV)](#export-객체--csv)
8. [Export 샘플 (클래스 → 샘플 CSV)](#export-샘플-클래스--샘플-csv)
9. [Import (엑셀 → 객체)](#import-엑셀--객체)
10. [Import (CSV → 객체)](#import-csv--객체)
11. [자주 묻는 질문](#자주-묻는-질문)
12. [빌드 & 기여](#빌드--기여)
13. [라이선스](#라이선스)

---

## 주요 기능

- **객체 → 엑셀과 엑셀 → 객체를 선언 하나로** — 같은 `@PxlColumn` 매핑을 exporter와 importer가 함께
  읽으므로 라운드트립이 어긋나지 않는다.
- **XLSX · XLS · 스트리밍 XLSX** — POI 엔진(`XSSF` / `HSSF` / `SXSSF`)을 애노테이션 속성 하나로 고르며,
  파일 형식 · 확장자 · 시트/행/열 한도가 거기서 따라온다.
- **CSV도 같은 모델로** — 같은 애노테이션, 같은 컨버터, 같은 컬럼 순서가 워크북 대신 CSV를 쓰고 그대로
  다시 읽는다.
- **선언만으로 매핑** — 필드에 `@PxlColumn`, 컬렉션 필드에 `@PxlSheet`, 클래스에 `@PxlWorkbook`. DTO의
  나머지는 건드리지 않는다. 헤더는 이름으로 매칭하므로 파일의 열 순서는 자유이고 정의에 없는 열은 무시된다.
- **약 30종 필드 타입 기본 지원** — 원시 타입과 래퍼, `String`, `BigDecimal`, `BigInteger`, `enum`,
  `LocalDate` / `LocalDateTime` / `LocalTime` / `Date`, `Duration` / `Period`, `UUID`, `Collection`, 그리고
  그 밖의 타입은 `@PxlExportConverter` / `@PxlImportConverter`로 처리한다.
- **행 단위 Bean Validation** — DTO에 선언한 표준 `javax.validation` / `jakarta.validation` 제약이 바인딩
  중에 검증되며, PXL 자체 제약 `@PxlByteSize`도 함께 쓸 수 있다.
- **엑셀 표현 요소** — 셀 스타일러, 컬럼 폭, 행 높이, 창 고정, 자동 필터, 드롭다운 검증, 행 그룹핑,
  이미지 삽입, 암호가 걸린 워크북.
- **대용량 파일 대응** — export는 SXSSF로 돌릴 수 있고 CSV export는 4 MiB를 넘어서면 힙 대신 임시
  파일로 이어 쓰며, import는 `excel-streaming-reader` 위에서 돌릴 수 있다.
- **클래스만으로 샘플 양식 생성** — 헤더 행 + 예시 값이 채워진 데이터 행 1개짜리 `.xlsx` · `.csv` 양식을
  만들어 배포하고 채워 받을 수 있다.
- **독립된 두 채널의 다국어 지원** — 콘텐츠 쪽으로 시트명 · 컬럼명을 번역하고, PXL 자체 진단 · 예외 메시지의
  언어를 따로 전환한다(영어 · 한국어 번들 동봉).
- **소스 하나에서 나오는 두 변형** — Java 8+용 `pxl-javax`와 Java 17+용 `pxl-jakarta`.

---

## 구성

환경에 맞는 변형 하나만 의존성에 추가한다.
- `pxl-javax`(Java 8+, `javax.*`)
- `pxl-jakarta`(Java 17+, `jakarta.*`)

### Maven

```xml
<!-- javax 변형 (Java 8+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-javax</artifactId>
    <version>0.9.4</version>
</dependency>
```

```xml
<!-- jakarta 변형 (Java 17+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-jakarta</artifactId>
    <version>0.9.4</version>
</dependency>
```

### Gradle

```groovy
// javax 변형 (Java 8+)
implementation 'io.github.hclimkr:pxl-javax:0.9.4'
```

```groovy
// jakarta 변형 (Java 17+)
implementation 'io.github.hclimkr:pxl-jakarta:0.9.4'
```

---

## 객체 DTO 정의

### 행 클래스

행 클래스는 `@PxlColumn`으로 각 필드를 헤더에 매핑한다.

```java
import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.time.LocalDate;

@Getter                     // (선택) 사용자 편의용 — PXL은 직접 바인딩하기 때문에 Getter가 필요하지는 않다.
@Setter                     // (선택) 사용자 편의용 — PXL은 직접 바인딩하기 때문에 Setter가 필요하지는 않다.
@NoArgsConstructor          // (필수) import 시에 무인자 생성자 필수
public class Employee {

    @PxlRowIndex            // (선택) 1-based 스프레드시트 행 번호. 타입: byte/short/int/long + 래퍼 클래스(Byte/Short/Integer/Long)
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

- `name`을 생략하면 필드명이 열 이름이 된다.
- `name`은 실제 헤더와 일치해야 바인딩된다(공백은 무시, 대소문자는 구분).
- `exportSample`은 [Export 샘플](#export-샘플-클래스--샘플-엑셀)에 들어갈 예시 값이다(일반 export에는 영향 없음).
- `exportSample`은 `String`으로 쓰지만 컬럼 타입으로 파싱되므로 그 타입이 받아들이는 값이어야 한다(아니면 `PxlCellCodecException`). 지정하지 않은 컬럼은 샘플 행에 `exportNullString`(기본 `""`)이 들어간다.

`Grade`는 예제에서 쓰는 사용자 정의 enum이다.

```java
public enum Grade {
    A, B, C, F
}
```

다중 시트 예제에서 두 번째 시트로 쓰는 행 클래스도 같은 방식으로 정의한다.

```java
import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter                     // (선택) 사용자 편의용 — PXL은 직접 바인딩하기 때문에 Getter가 필요하지는 않다.
@Setter                     // (선택) 사용자 편의용 — PXL은 직접 바인딩하기 때문에 Setter가 필요하지는 않다.
@NoArgsConstructor          // (필수) import 시에 무인자 생성자 필수
public class Department {

    @PxlColumn(name = "Code")
    private String code;

    @PxlColumn(name = "DepartmentName")
    private String departmentName;

    @PxlColumn(name = "Headcount")
    private int headcount;
}
```

### 워크북 클래스 (다중 시트를 한 객체로)

각 시트 필드는 `Collection` 타입이고 `@PxlSheet`로 바인딩한다.

```java
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter                     // (선택) 사용자 편의용 — PXL은 직접 바인딩하기 때문에 Getter가 필요하지는 않다.
@Setter                     // (선택) 사용자 편의용 — PXL은 직접 바인딩하기 때문에 Setter가 필요하지는 않다.
@NoArgsConstructor          // (필수) import 시에 무인자 생성자 필수
public class Company {

    @PxlWorkbookName        // (선택) 워크북 이름을 담을 String 필드
    private String workbookName;

    @PxlSheet(name = "Employees")
    private List<Employee> employees;

    @PxlSheet(name = "Departments")
    private List<Department> departments;
}
```

---

## 한눈에 보는 사용법

### Export

```java
import io.github.hclimkr.pxl.Pxl;

import java.io.File;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

// Pxl은 한 번 만들어 재사용한다 (thread-safe, 상태 비공유 — 싱글톤/스프링 빈 권장)
Pxl pxl = new Pxl();

// Employee 행 객체 준비
Employee alice = new Employee();
alice.setName("Alice");
alice.setAge(30);
alice.setSalary(50_000L);
alice.setActive(true);
alice.setHireDate(LocalDate.of(2020, 1, 15));
alice.setGrade(Grade.A);

// Employee 행 객체 준비
Employee bob = new Employee();
bob.setName("Bob");
bob.setAge(42);
bob.setSalary(72_000L);
bob.setActive(false);
bob.setHireDate(LocalDate.of(2018, 6, 1));
bob.setGrade(Grade.B);

// Employees 시트 객체 준비
List<Employee> employees = Arrays.asList(alice, bob);

// Department 행 객체 준비
Department eng = new Department();
eng.setCode("ENG");
eng.setDepartmentName("Engineering");
eng.setHeadcount(12);

// Department 행 객체 준비
Department sal = new Department();
sal.setCode("SAL");
sal.setDepartmentName("Sales");
sal.setHeadcount(8);

// Departments 시트 객체 준비
List<Department> departments = Arrays.asList(eng, sal);

// Company 워크북 객체 준비
Company company = new Company();
company.setEmployees(employees);
company.setDepartments(departments);

// Export: Company 워크북 객체 → 엑셀 파일
pxl.exportExcel()
   .workbook(company)
   .toFile(new File("company.xlsx"));
```

### Import

```java
import io.github.hclimkr.pxl.Pxl;

import java.io.File;

// Pxl은 한 번 만들어 재사용한다 (thread-safe, 상태 비공유 — 싱글톤/스프링 빈 권장)
Pxl pxl = new Pxl();

// Import: 엑셀 파일 → Company 워크북 객체
Company company = pxl.importExcel()
                     .workbook(Company.class)
                     .fromFile(new File("company.xlsx"));
```
모든 작업은 위 예제처럼 하나의 메서드 체인으로 처리한다. 맨 앞 메서드 이름이 작업의 방향(내보내기/가져오기)과 형식(엑셀/CSV)을 나타내며, 이어서 대상을 지정한 뒤 마지막 메서드에서 실행된다.

| 용도           | 메서드 체인 (시작 → 구성 → 실행)                                                                                                                                                                    |
|--------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 엑셀 export    | `pxl.exportExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                                     |
| 샘플 엑셀 export | `pxl.exportSampleExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                               |
| CSV export   | `pxl.exportCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                          |
| 샘플 CSV export | `pxl.exportSampleCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                    |
| 엑셀 import    | `pxl.importExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromStream(InputStream)`                                                                                    |
| CSV import   | `pxl.importCsv()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromFiles(List<File>)` / `.fromStream(String, InputStream)` / `.fromStreams(List<String>, List<InputStream>)` |

---

## Export (객체 → 엑셀)

`workbook(...)` 또는 `sheet(...)`로 내용을 구성하고(두 형태 중 하나만 지정하며, 함께 지정하면 예외가 발생한다),
마지막에 `toFile(File)` / `toStream(OutputStream)` / `toWorkbook()` 중 하나로 출력한다.

### 워크북 객체 → 엑셀

```java
Company company = ...;

pxl.exportExcel()
   .workbook(company)
   .toFile(new File("company.xlsx"));
```

### 단일 시트 객체 → 엑셀

```java
pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .toFile(new File("employees.xlsx"));
```

### Export 결과 엑셀 모양 (단일 시트)

위 `Employee` 리스트를 `sheet(Employee.class, employees, "Employees")`로 export하면 "Employees" 시트가 다음처럼 만들어진다(열 문자 A~F·행 번호는 엑셀 화면 기준).

|       | A     | B   | C      | D      | E          | F     |
|-------|-------|-----|--------|--------|------------|-------|
| **1** | Name  | Age | Salary | Active | HireDate   | Grade |
| **2** | Alice | 30  | 50000  | true   | 2020-01-15 | A     |
| **3** | Bob   | 42  | 72000  | false  | 2018-06-01 | B     |

- 1행은 헤더행이고, 2행부터는 데이터행이다.
- 셀 타입: `Name`은 텍스트, `Age`·`Salary`는 숫자, `Active`는 `true`/`false` 텍스트, `HireDate`는 텍스트, `Grade`는 텍스트다
- `@PxlRowIndex`(`rowIndex`)는 열이 아니므로 시트에 기록되지 않는다.

### 다중 시트 객체 형태

```java
pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .sheet(Department.class, departments, "Departments")
   .toFile(new File("company.xlsx"));
```

### Export 결과 엑셀 모양 (다중 시트)

워크북 객체(`Company`)나 `.sheet(...)` 반복 호출로 export하면 한 `.xlsx`에 시트가 여러 개 생긴다 — `@PxlSheet` 필드(또는 각 `.sheet(...)`) 하나가 시트 하나다. 위 예제는 시트 탭 `[ Employees | Departments ]` 두 개를 만든다.

- "Employees" 시트 — 위 [Export 결과 엑셀 모양 (단일 시트)](#export-결과-엑셀-모양-단일-시트)의 `Employee` 표와 동일.
- "Departments" 시트가 다음처럼 만들어진다.

|       | A    | B              | C         |
|-------|------|----------------|-----------|
| **1** | Code | DepartmentName | Headcount |
| **2** | ENG  | Engineering    | 12        |
| **3** | SAL  | Sales          | 8         |

- 워크북 클래스의 `@PxlWorkbookName` `String` 필드는 열(셀)이 아니라 워크북 이름 용도이므로 어느 시트에도 나타나지 않는다.
- Import는 대칭이다 — 이 두 시트를 가진 `.xlsx`를 `workbook(Company.class)`로 읽으면 시트명이 매칭되어 각 `@PxlSheet` 필드(`employees`, `departments`)가 채워진다.

### 출력 대상

```java
// 1) 파일 (PXL 내부에서 파일을 열고 닫는다)
pxl.exportExcel()
   .sheet(Row.class, rows, "S")
   .toFile(new File("out.xlsx"));
```

```java
// 2) 스트림 (HTTP 응답 등) — PXL는 스트림을 닫지 않고 호출자가 닫는다
try (OutputStream os = response.getOutputStream()) {
    pxl.exportExcel()
       .sheet(Row.class, rows, "S")
       .toStream(os);
}
```

```java
// 3) POI Workbook — 반환된 워크북은 호출자가 닫는다
Workbook workbook = pxl.exportExcel()
                       .sheet(Row.class, rows, "S")
                       .toWorkbook();
```

---

## Export 샘플 (클래스 → 샘플 엑셀)

클래스만으로 헤더행 + 각 컬럼 예시 값(`exportSample`)으로 채운 데이터행 1개짜리 샘플 양식을 만든다.
마지막 출력 메서드(`toFile`/`toStream`/`toWorkbook`)는 일반 export와 동일하다.

```java
// 워크북 클래스 형태
pxl.exportSampleExcel()
   .workbook(Company.class)
   .toFile(new File("sample.xlsx"));
```

```java
// 단일 시트 형태
pxl.exportSampleExcel()
   .sheet(Employee.class, "Employees")
   .toFile(new File("sample.xlsx"));
```

```java
// 다중 시트 형태
pxl.exportSampleExcel()
   .sheet(Employee.class, "Employees")
   .sheet(Department.class, "Departments")
   .toFile(new File("sample.xlsx"));
```

### Export 샘플 결과 엑셀 모양

위 `Employee`로 `exportSampleExcel().sheet(Employee.class, "Employees")`를 만들면 헤더 행 + 샘플 데이터 행 1개짜리 "Employees" 시트가 생긴다.

|       | A        | B   | C      | D      | E          | F     |
|-------|----------|-----|--------|--------|------------|-------|
| **1** | Name     | Age | Salary | Active | HireDate   | Grade |
| **2** | John Doe | 25  | 45000  | true   | 2024-03-01 | C     |

- 다중 시트(`Company` 또는 `.sheet(...).sheet(...)`)면 시트마다 같은 방식의 헤더 행 + 샘플 데이터 행 1개짜리가 만들어진다.

---

## Export (객체 → CSV)

CSV 파일 하나에는 시트 하나가 담기므로 이 빌더는 시트 형태만 있고(`workbook(...)`이 없다) 마지막 실행 메서드가
그 시트 하나를 쓴다. 값을 만드는 부분은 엑셀 export와 전부 공유한다 — 같은 애노테이션, 같은 컨버터, 같은 컬럼 순서다.

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

### 인코딩 · 구분자 · BOM

CSV 워크북은 한 시트가 곧 한 파일이므로 이 셋은 스키마가 아니라 파일의 속성이고, 시트 단계에서 정해진다.
시트 형태에는 `@PxlSheet`를 붙일 필드가 없으므로 와일드카드 `PxlExportSheetOption`이 시트 단계 경로이고,
워크북 단계 옵션은 모든 시트에 한 번에 적용된다.

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

- BOM은 UTF-8 · UTF-16LE · UTF-16BE에서만 기록된다. 그 외 문자셋에서는 조용히 생략된다 — UTF-16은 인코더가
  이미 하나를 붙이고, EUC-KR 같은 비유니코드 문자셋은 아예 인코딩할 수 없어 첫 필드가 깨지기 때문이다.
- 헤더 줄은 항상 기록된다. 끄는 속성은 없다.

### CSV가 담지 못하는 것

셀의 겉모습이나 워크북의 구성을 정하는 설정은 무시된다 — 스타일러, 컬럼 폭, 행 높이, 창 고정, 자동 필터,
드롭다운, 엔진과 스트리밍 윈도우, 그리고 `exportGroupingFieldName`(행은 넘긴 순서 그대로 나간다)이다.
다만 둘은 값이 파일에 그대로 남는다 — `exportStringAsFormula`는 앞의 `=`까지 포함해 텍스트가 그대로 기록되고,
`exportStringAsPicture`는 그림을 박는 대신 이미지 경로가 필드 값이 된다.

`exportPassword`만은 무시가 아니라 거부된다. CSV는 암호화할 수 없는데 대신 평문을 쓰면 유출이기 때문이다.

---

## Export 샘플 (클래스 → 샘플 CSV)

클래스만으로 헤더 1줄 + 각 컬럼 예시 값(`exportSample`)으로 채운 데이터 1줄짜리 샘플 양식을 만든다 —
[Export 샘플 (클래스 → 샘플 엑셀)](#export-샘플-클래스--샘플-엑셀)의 CSV판이다.
CSV 파일 하나에는 시트 하나가 담기므로 이 빌더도 시트 형태만 있고, 마지막 실행 메서드는 `toFile`/`toStream`이다
(`toWorkbook`은 없다).

```java
pxl.exportSampleCsv()
   .sheet(Employee.class, "Employees")
   .toFile(new File("employees-template.csv"));
```

### Export 샘플 결과 CSV 모양

위 `Employee`로 `exportSampleCsv().sheet(Employee.class, "Employees")`를 만들면 다음 파일이 나온다.

```csv
Name,Age,Salary,Active,HireDate,Grade
John Doe,25,45000,true,2024-03-01,C
```

- `exportSample`이 없는 컬럼은 `exportNullString`(기본 `""`)이 들어가므로 해당 필드는 빈 값이 된다.
- `importCsv()`로 그대로 다시 읽히므로, 배포해서 채워 받는 양식으로 쓸 수 있다.
- CSV 파일 자체에는 시트 이름이 담기지 않는다. 여기 넘긴 이름은 시트 단계 옵션을 고르고 오류 메시지에 표시되는 용도다.
- 인코딩 · 구분자 · BOM은 [Export (객체 → CSV)](#export-객체--csv)와 동일하게 정해지고, 무시·거부되는 설정도 같다.

---

## Import (엑셀 → 객체)

`importExcel()`로 시작해 워크북 형태(`workbook(Class)`) 또는 시트 형태(`sheet(Class, 후보시트명...)`)로 읽을 대상을 구성한 뒤,
마지막 실행 메서드(`fromFile`/`fromStream`)를 지정하면 그 자리에서 파싱된다.

### Import 대상 엑셀 모양

읽을 "Employees" 시트가 아래 모양이면 `Employee`로 바인딩된다(위 [Export 결과 엑셀 모양](#export-결과-엑셀-모양-단일-시트)과 같은 구조 — 라운드트립).

|       | A     | B   | C      | D      | E          | F     |
|-------|-------|-----|--------|--------|------------|-------|
| **1** | Name  | Age | Salary | Active | HireDate   | Grade |
| **2** | Alice | 30  | 50000  | true   | 2020-01-15 | A     |
| **3** | Bob   | 42  | 72000  | false  | 2018-06-01 | B     |

- 헤더 이름이 `@PxlColumn(name=...)`과 일치해야 매칭된다(공백은 무시, 대소문자는 구분). 열 순서는 자유이고, 정의에 없는 열은 무시된다.
- `@PxlRowIndex`(`rowIndex`)는 읽은 각 행의 1-based 스프레드시트 행 번호로 자동 채워진다.

### 워크북 형태 (@PxlWorkbook 객체)

```java
Company company = pxl.importExcel()
                     .workbookName("Acme")            // (선택) @PxlWorkbookName 필드에 채울 값
                     .workbook(Company.class)
                     .fromFile(new File("company.xlsx"));
```

### 시트 형태 (Collection 반환)

```java
// 파일에서 List로 반환
List<Employee> rows = pxl.importExcel()
                         .sheet(Employee.class, "Employees")
                         .fromFile(new File("employees.xlsx"));
```

```java
// 파일에서 Set로 반환
Set<Employee> set = pxl.importExcel()
                       .sheet(Employee.class, Set.class, "Employees")
                       .fromFile(new File("employees.xlsx"));
```

```java
// 스트림에서 List로 반환 (PXL는 스트림을 닫지 않고 호출자가 닫는다)
try (InputStream is = new FileInputStream("employees.xlsx")) {
    List<Employee> rows2 = pxl.importExcel()
                              .sheet(Employee.class, "Employees")
                              .fromStream(is);
}
```

---

## Import (CSV → 객체)

엑셀 import와 같은 방식이다(시작 → 워크북/시트 구성 → 마지막 실행 메서드 호출). CSV는 파일명(확장자 제외)이 시트명이 되고, 마지막 실행 메서드가 `fromFile`/`fromFiles`/`fromStream`/`fromStreams`로 늘어난다.
시트 형태는 후보 시트명 인자 없이 `sheet(Class)`(단일 CSV), 워크북 형태는 여러 CSV를 시트별로 묶는다.

### Import 대상 CSV 파일 모양

CSV는 첫 줄이 헤더, 이후가 데이터인 일반 텍스트다(엑셀과 달리 셀 타입이 없어 날짜·불리언도 문자열로 적는다).
파일명에서 확장자를 뺀 이름이 시트명이 된다. 워크북 형태에서는 이 이름이 `@PxlSheet` 시트명과 일치해야 바인딩되며, 공백과 대소문자는 무시한다.

```text
Name,Age,Salary,Active,HireDate,Grade
Alice,30,50000,true,2020-01-15,A
Bob,42,72000,false,2018-06-01,B
```

위 `Employees.csv`를 `sheet(Employee.class).fromFile(...)`로 읽으면 아래 `List<Employee>`가 된다.

| Name  | Age | Salary | Active | HireDate   | Grade |
|-------|-----|--------|--------|------------|-------|
| Alice | 30  | 50000  | true   | 2020-01-15 | A     |
| Bob   | 42  | 72000  | false  | 2018-06-01 | B     |

- 기본 인코딩은 `UTF-8`, 구분자는 `,`다.
- 헤더 이름 매칭·값 해석(불리언 토큰, 열 순서 자유, 정의에 없는 열 무시 등)은 엑셀 import와 동일하다(위 [Import 대상 엑셀 모양](#import-대상-엑셀-모양) 참고).
- 워크북 형태로 여러 CSV를 묶을 때는 각 파일이 한 시트가 된다 — 예: `Departments.csv` → 시트 `Departments`.

### 워크북 형태 (여러 CSV를 시트별로 묶기)

```java
// 파일 여러 개 → 각 파일이 하나의 시트
Company company = pxl.importCsv()
                     .workbookName("Acme")
                     .workbook(Company.class)
                     .fromFiles(Arrays.asList(employeesCsv, departmentsCsv));
```

```java
// 스트림 여러 개 (이름 리스트 + 스트림 리스트)
Company company = pxl.importCsv()
                     .workbook(Company.class)
                     .fromStreams(names, streams);
```

### 시트 형태 (단일 CSV)

```java
// 파일 (시트명 인자 없음 — CSV는 단일 표)
List<Employee> employees = pxl.importCsv()
                              .sheet(Employee.class)
                              .fromFile(new File("Employees.csv"));
```

```java
// 스트림 (이름 + 스트림)
List<Employee> employees = pxl.importCsv()
                              .sheet(Employee.class)
                              .fromStream("Employees", inputStream);
```

---

더 자세한 내용(타입별 동작, 전체 애노테이션 속성, i18n, 스타일러, 예외 등)은 [docs/reference_ko.md](docs/reference_ko.md) 를 참고한다.

---

## 자주 묻는 질문

**엑셀 파일을 자바 객체 리스트로 어떻게 읽나?**
DTO 필드에 `@PxlColumn`을 붙이고
`pxl.importExcel().sheet(Employee.class, "Employees").fromFile(file)`을 호출한다. 모든 셀이 필드 타입으로
변환된 `List<Employee>`가 반환된다 — [Import (엑셀 → 객체)](#import-엑셀--객체) 참고.

**스프링 컨트롤러에서 엑셀 다운로드로 어떻게 내려주나?**
`toStream(...)`으로 서블릿 출력 스트림에 쓴다. PXL은 스트림을 닫지 않으므로 응답의 제어권은 호출자가
그대로 쥔다 — [출력 대상](#출력-대상) 참고. 응답에 실을 content type과 확장자는 `PxlFileFormat`에서 얻는다.

**메모리에 다 올릴 수 없는 대용량 파일도 되나?**
export는 `SXSSF` 엔진을 쓸 수 있고, CSV export는 4 MiB를 넘으면 임시 파일로 넘어간다. import는
`@PxlWorkbook(importUsingStreamReader = true)`로 `excel-streaming-reader` 위에서 돌릴 수 있다(XLSX
전용이며 수식 셀은 평가하지 못한다).

**`.xlsx` 말고 `.xls`도 읽나?**
읽는다. export는 `@PxlWorkbook(exportExcelEngine = PxlExcelEngine.HSSF)`로 `.xls`를 고르고,
import는 파일 자체에서 형식을 탐지한다.

**CSV도 지원하나?**
`exportCsv()` / `importCsv()`로 지원하며 애노테이션은 엑셀과 동일하다. 문자셋 · 구분자 · BOM을 지정할 수
있고, 여러 CSV 파일을 파일 하나당 시트 하나로 묶어 한 워크북처럼 읽을 수 있다.

**PXL이 모르는 타입의 컬럼은 어떻게 하나?**
클래스에 `@PxlExportConverter` / `@PxlImportConverter` 메서드 쌍을 두면 해당 컬럼이 그 변환기를 거친다.

**`pxl-javax`와 `pxl-jakarta` 중 무엇을 쓰나?**
Java 8 이상에서 `javax.validation`을 쓰면 `pxl-javax`, Java 17 이상에서 `jakarta.validation`을 쓰면
`pxl-jakarta`다. 같은 라이브러리이므로 정확히 하나만 추가한다.

**Bean Validation이 반드시 있어야 하나?**
아니다. 구현체가 클래스패스에 없으면 PXL은 경고 로그만 남기고 검증을 건너뛴 채 그대로 동작한다.

**`Pxl` 인스턴스를 공유해도 되나?**
된다 — 상태가 없고 thread-safe하므로 한 번 만들어 싱글톤이나 스프링 빈으로 재사용한다.

---

## 빌드 & 기여

소스 코드는 `pxl-javax`에만 있고 `pxl-jakarta`는 빌드 시 문자열 치환으로 생성된다.  
이 저장소는 이슈 보고와 제안만 받는다 — [CONTRIBUTING_ko.md](CONTRIBUTING_ko.md) 를 참고한다.

---

## 라이선스

이 프로젝트는 [Apache License 2.0](LICENSE) 하에 배포된다.

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
