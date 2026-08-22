[English](reference.md) · **한국어**

PXL 레퍼런스
=============================

[![빠른 시작](https://img.shields.io/badge/🚀%20빠른%20시작-README_ko.md-4c9aff?style=for-the-badge)](../README_ko.md)

PXL은 **애노테이션 기반으로 스프레드시트와 자바 객체를 양방향 바인딩**하는 라이브러리이다.
Apache POI와 Apache Commons CSV 위에 구축되었으며, Java 8 이상을 지원한다.

- 내부적으로 Apache POI로 엑셀(XLS/XLSX)을, Apache Commons CSV로 CSV를 처리한다.
- Import: XLS, XLSX, CSV → 자바 객체
- Export: 자바 객체 → 엑셀 (기본 XLSX, `@PxlWorkbook(exportExcelEngine = ...)`로 XLS·스트리밍 XLSX 선택 가능)
- 애노테이션이 붙은 필드/클래스만 바인딩 대상이 된다.

> 🚀 **처음이라면 → [README_ko.md](../README_ko.md)** — 예제 위주로 가장 빠르게 적용하는 실전 가이드.
> 이 문서는 지원 변수 타입·전체 옵션·제약까지 다루는 레퍼런스다.

## 목차

1. [왜 PXL인가](#왜-pxl인가)
2. [구성](#구성)
3. [API 구조](#api-구조)
4. [지원 변수 타입](#지원-변수-타입)
5. [애노테이션](#애노테이션)
6. [옵션 오버라이드](#옵션-오버라이드)
7. [유효성 검사](#유효성-검사)
8. [예외](#예외)
9. [i18n](#i18n)
10. [시트 내의 행/열 인덱스 규칙](#시트-내의-행열-인덱스-규칙)
11. [셀 스타일러](#셀-스타일러)
12. [다양한 예제](#다양한-예제)
13. [제약 (Limitation)](#제약-limitation)
14. [자주 겪는 함정 체크리스트](#자주-겪는-함정-체크리스트)
15. [라이선스](#라이선스)

---

## 왜 PXL인가

애노테이션만으로 스프레드시트와 자바 객체 간 매핑을 가장 정확하고 폭넓게 다루는 것을 목표로 한다.
DTO에 애노테이션을 붙이는 선언적 방식으로 아래 기능을 별도 코드 없이 얻는다.

- **타입 충실성 & 엄격성**   
  완전한 `java.time`(`Zoned`/`Offset`/`Duration`/`Period` 포함), `BigInteger`/`BigDecimal` 정밀도(2^53) 인지,
  `NaN`/`Infinity` 거부, non-lenient 날짜 파싱(무효 날짜 rollover 차단), `Collection` 위치 보존,
  `UUID`의 canonical 형식 강제, export/import 대칭 동작.
  흔한 타입만이 아니라 이런 엣지까지 코덱 단위로 방어·문서화한다.
- **표준 유효성 검사 통합 + 커스텀 제약**  
  import 시 `javax.validation`/`jakarta.validation`으로 행 단위 유효성 검사하고,
  `@PxlByteSize`(바이트 길이) 같은 커스텀 제약을 별도로 제공한다.
- **헤더 i18n**  
  시트·컬럼명을 `ResourceBundle`로 번역해 다국어 템플릿을 만든다.
- **javax + jakarta 동시 지원**  
  단일 소스 코드에서 두 아티팩트를 생성해 레거시(`javax`)와 신규(`jakarta`) 환경을 모두 커버한다.
- **애노테이션 기반 샘플/양식 export**  
  애노테이션이 붙은 클래스로부터 입력 양식(드롭다운·샘플값·i18n 헤더 포함)을 바로 생성한다.
- **폭넓은 POI 기능을 애노테이션으로**  
  드롭다운/데이터검증, 이미지 삽입(URL 포함), 수식(`exportStringAsFormula`), 암호화,
  필드값 별 시트 분할, 스타일러 캐스케이드, auto-size.
- **하나의 애노테이션 모델로 다중 포맷**  
  읽기: XLS/XLSX/CSV, 쓰기: XLSX/XLS/스트리밍 XLSX(SXSSF).

---

## 구성

환경에 맞는 변형 하나만 의존성에 추가한다.

**Maven**

```xml
<!-- javax 변형 (Java 8+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-javax</artifactId>
    <version>0.9.5</version>
</dependency>
```

```xml
<!-- jakarta 변형 (Java 17+) -->
<dependency>
    <groupId>io.github.hclimkr</groupId>
    <artifactId>pxl-jakarta</artifactId>
    <version>0.9.5</version>
</dependency>
```

**Gradle**

```groovy
// javax 변형 (Java 8+)
implementation 'io.github.hclimkr:pxl-javax:0.9.5'
```

```groovy
// jakarta 변형 (Java 17+)
implementation 'io.github.hclimkr:pxl-jakarta:0.9.5'
```

## 런타임 의존성

Bean Validation API(`validation-api` / `jakarta.validation-api`)와
로깅 파사드 `slf4j-api`는 PXL가 `compile` 스코프로 선언해 자동 전이되므로
사용자가 따로 넣지 않아도 된다.

아래 두 가지만 필요할 때 사용자가 직접 넣는다.

### Bean Validation 구현체 + EL  

`@NotNull`·`@NotEmpty`·`@NotBlank` 등에 대해 데이터 유효성 검사를 하려면, 구현체(예: hibernate-validator)와 EL(예: jakarta.el)을 사용자가 추가한다.  
둘 중 하나라도 없으면 `new Pxl()`은 경고 로그(SLF4J `WARN`)만 남긴 채 유효성 검사를 비활성화한다.  

**Maven 예시**

```xml
<!-- javax 변형 -->
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
<!-- jakarta 변형 -->
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

**Gradle 예시**

```groovy
// javax 변형
implementation 'org.hibernate.validator:hibernate-validator:6.2.5.Final'
implementation 'org.glassfish:jakarta.el:3.0.4'
```

```groovy
// jakarta 변형
implementation 'org.hibernate.validator:hibernate-validator:9.1.1.Final'
implementation 'org.glassfish:jakarta.el:4.0.2'
```

### SLF4J 바인딩

로그를 실제로 출력하려면 SLF4J 바인딩(`logback-classic`·`slf4j-simple`·`log4j-slf4j2-impl` 등)을 사용자가 추가한다.  
바인딩이 없으면 SLF4J는 `No SLF4J providers were found` 경고 후 로그를 버리도록(NOP) 동작한다.  
POI 코어 로깅은 `log4j-api`이며 POI가 전이 제공하므로 별도 추가가 필요 없다.

**Maven 예시**

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

**Gradle 예시**

```groovy
implementation 'org.apache.logging.log4j:log4j-slf4j2-impl:2.26.1'
implementation 'org.apache.logging.log4j:log4j-core:2.26.1'
```

---

## API 구조

### 메서드 체인

각 작업은 시작 메서드부터 마지막 메서드까지 하나의 메서드 체인으로 수행한다.  
방향(export/import)과 형식(excel/csv)은 시작 메서드 이름에 담겨 있다.

| 용도           | 메서드 체인 (시작 → 구성 → 실행)                                                                                                                                                                    |
|--------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 엑셀 export    | `pxl.exportExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                                     |
| 샘플 엑셀 export | `pxl.exportSampleExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)` / `.toWorkbook()`                                                               |
| CSV export   | `pxl.exportCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                          |
| 샘플 CSV export | `pxl.exportSampleCsv()`<br/>→ `.sheet(...)`<br/>→ `.toFile(File)` / `.toStream(OutputStream)`                                                                                                    |
| 엑셀 import    | `pxl.importExcel()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromStream(InputStream)`                                                                                    |
| CSV import   | `pxl.importCsv()`<br/>→ `.workbook(...) / .sheet(...)`<br/>→ `.fromFile(File)` / `.fromFiles(List<File>)` / `.fromStream(String, InputStream)` / `.fromStreams(List<String>, List<InputStream>)` |

- 구성 단계의 `.workbook(...)`과 `.sheet(...)`은 서로 배타적이다 — 한 체인에서 둘을 함께 지정하면 `PxlArgumentException`이 발생한다(둘 다 생략해도 같다).
- CSV export 체인 2종에는 `.workbook(...)` 자체가 없다 — CSV 파일 하나에는 시트 하나가 담기므로 워크북 형태를 호출할 자리가 없다. `.sheet(...)`은 엑셀과 마찬가지로 누적되지만 마지막 메서드가 시트 하나를 쓰므로, 시트를 둘 이상 구성하면 구성 단계가 아니라 마지막 메서드에서 `PxlArgumentException`(`builder.export.csv.singleSheetOnly`)이 발생한다.
- export는 `.sheet(...)`를 여러 번 호출해 여러 시트를 만든다 — 호출 순서가 그대로 워크북 내 시트 순서가 되며, 시트마다 다른 행 클래스를 줄 수 있다. 시트 이름이 중복되면(안전한 이름으로 정규화한 뒤 대소문자를 무시하고 비교) 겹친 이름을 담은 `PxlDataException`이 발생한다 — 대소문자만 다른 두 시트는 한 워크북에 공존할 수 없다.
- import는 `.sheet(...)`를 연달아 체인할 수 없다 — 여러 시트를 읽는 방법은 두 가지다.
    - 워크북 형태로 한 번에: `@PxlWorkbook` 클래스를 `.workbook(...)`에 주면 `@PxlSheet` 필드별로 여러 시트가 한 번에 바인딩된다.  
      CSV는 한 파일이 한 시트이므로 이 형태에서 소스를 `.fromFiles(List<File>)`·`.fromStreams(List<String>, List<InputStream>)`로 여러 개 넘긴다(파일명에서 확장자를 뗀 이름·`csvNames`가 `@PxlSheet` 이름과 매칭된다 — 공백과 대소문자를 무시한다).
    - 시트별로 나눠서: 같은 빌더 인스턴스로 `.sheet(...)`를 시트마다 호출하고 각각 마지막(실행) 단계까지 실행한다. 소스는 실행 단계에서 지정하므로 스트림은 호출마다 새로 열어 준다(파일은 매번 내부에서 열고 닫는다).
- 구성 단계의 `.override(...)`은 선택적이며 체인 안에서 위치를 자유롭게 정할 수 있다 — `.workbook(...)`/`.sheet(...)`의 앞이든 뒤든 마지막(실행) 단계 전이기만 하면 된다(여러 번 지정하면 마지막 값이 적용된다). 옵션 객체에 담긴 값으로 애노테이션 값을 런타임에 오버라이드한다.  
  export는 `.override(...)`에 `PxlExportWorkbookOption`을, import는 `PxlImportWorkbookOption`을 인자로 넘긴다(생략하면 애노테이션 값을 그대로 쓴다).
  import는 워크북 이름을 덮어쓰는 `.workbookName(String)`도 같은 위치에 둘 수 있다. 생략하면 파일에서 읽는 엑셀 import는 확장자를 제거한 파일명을 워크북 이름으로 쓴다 — [`@PxlWorkbookName`](#pxlworkbookname-필드-대상) 참조.
- 각 옵션의 필드 목록과 빌더 예시는 [옵션 오버라이드](#옵션-오버라이드) 절을 참고한다.

### 자원 소유권

| 마지막 메서드                                                                             | 파라미터                | PXL의 처리                      |
|-------------------------------------------------------------------------------------|---------------------|------------------------------|
| export<br/>`toFile(File)`                                                                | 호출자의 `File`          | 내부에서 파일의 스트림을 열고 직접 닫는다      |
| export<br/>`toStream(OutputStream)`                                                     | 호출자의 `OutputStream` | 닫지 않는다(flush만). 호출자가 닫아야 한다  |
| export<br/>`toWorkbook()`                                                               | 반환 `Workbook`       | 반환한 워크북을 닫지 않는다. 호출자가 닫아야 한다 |
| import<br/>`fromFile(File)`<br/>`fromFiles(List<File>)` (CSV)                                   | 호출자의 `File`                  | 내부에서 파일의 스트림을 열고 직접 닫는다      |
| import<br/>`fromStream(InputStream)` (Excel)<br/>`fromStream(String, InputStream)`·`fromStreams(List<String>, List<InputStream>)` (CSV) | 호출자의 `InputStream`  | 닫지 않는다. 호출자가 닫아야 한다          |

CSV export는 목적지를 열기 전에 출력물 전체를 만든다. 그래서 export가 실패하면 `toFile(...)`은 파일을 아예 만들지 않고,
`toStream(...)`은 호출자의 스트림에 한 번에 쓴다. 작은 출력은 메모리에 두고, 큰 출력은 임시 파일로 이어 쓴 뒤 호출이
끝나기 전에 그 파일을 지운다 — [제한사항](#제한사항) 참조.

### 빌더 수명과 스레드 안전성

`Pxl` 자체는 상태가 없고 스레드에 안전하다. 반면 `Pxl`이 만들어 주는 빌더는 그때까지 모은 구성을 들고 있으므로 스레드에 안전하지 않다.

| 객체                                              | 재사용            | 비고                                                                    |
|-------------------------------------------------|----------------|-----------------------------------------------------------------------|
| `Pxl`                                           | 재사용한다          | 상태 없음·스레드 안전. `new Pxl()`은 유효성 검사 부트스트랩 비용이 있으니 싱글톤으로 둔다 |
| export 빌더<br/>(`exportExcel()`·`exportSampleExcel()`) | 재실행은 되나 재구성은 안 된다 | 구성이 빌더에 남으므로 마지막(실행) 단계를 다시 호출하면 그대로 한 번 더 만들지만, 시트를 더 넣으면 기존 것 위에 누적된다 |
| CSV export 빌더<br/>(`exportCsv()`·`exportSampleCsv()`) | 재실행은 되나 재구성은 안 된다 | 위와 같고 `sheet(...)`이 누적되는 것도 같다 — 다만 CSV의 마지막 메서드는 시트 하나를 쓰므로, 두 번째 시트는 구성 단계가 아니라 마지막 메서드에서 실패한다 |
| import 빌더<br/>(`importExcel()`·`importCsv()`)   | 재사용할 수 있다      | 소스 스텝이 필요한 설정을 복사해 가므로 같은 빌더로 시트마다 실행할 수 있다                           |

export 빌더는 추가된 시트를 모두 보관하며, 어떤 마지막(실행) 단계도 이를 비우지 않는다.

덕분에 같은 구성으로 다시 실행하는 것은 정의된 동작이다 — 마지막(실행) 단계마다 빌더가 들고 있는 시트로 워크북을 새로 만들므로, 같은 내용을 여러 대상에 보낼 수 있다.

```java
PxlExcelExportBuilder builder = pxl.exportExcel()
                                   .sheet(Employee.class, employees, "Employees");

builder.toFile(new File("employees.xlsx"));   // 기록된다
builder.toStream(outputStream);               // 다시 만들어져 같은 내용이 나간다
```

이때 워크북은 매번 처음부터 다시 만들어진다 — 캐시된 결과를 재사용하는 것이 아니라 반복 생성이다.

동작하지 않는 것은 같은 빌더를 재구성하는 경우다. 실행 후 `sheet(...)`를 다시 호출하면 이미 있는 시트 위에 더해지므로 다음 실행에서 앞서 넣은 시트까지 함께 기록되고, 시트 이름이 겹치면 `PxlDataException`이 발생한다. 내용이 다른 워크북을 만들 때는 `exportExcel()`부터 체인을 새로 시작한다.

import 빌더는 반대다. `workbook(...)`/`sheet(...)`가 설정을 소스 스텝에 넘기고 빌더 자신은 그대로 남으므로, 재구성이야말로 이 빌더의 쓰임새다 — 빌더 하나로 시트를 얼마든지 읽을 수 있다.

```java
PxlExcelImportBuilder builder = pxl.importExcel();

List<Employee> employees = builder.sheet(Employee.class, "Employees")
                                  .fromFile(excelFile);
List<Department> departments = builder.sheet(Department.class, "Departments")
                                      .fromFile(excelFile);
```

호출마다 자기 시트를 읽으며, 반환 타입도 시트별로 다를 수 있다. 빌더에 지정한 `workbookName(...)`/`override(...)`는 그 뒤의 모든 실행에 적용된다. 특정 실행에만 적용하려면 소스 스텝에 지정한다 — 스텝의 값은 빌더에서 복사해 온 값과 병합되지 않고 그것을 대체한다.

---

### Export 엔진과 파일 형식

두 enum은 서로 다른 축에 있으며 어느 쪽도 다른 쪽을 대신하지 못한다.

| 타입               | 답하는 질문                | 상수                        |
|------------------|-----------------------|---------------------------|
| `PxlExcelEngine` | 어떤 POI 구현이 워크북을 쓰는가   | `HSSF` · `XSSF` · `SXSSF` |
| `PxlFileFormat`  | 나온 바이트가 무엇인가          | `XLS` · `XLSX` · `CSV`    |

`@PxlWorkbook(exportExcelEngine = ...)`와 `PxlExportWorkbookOption.exportExcelEngine`은 엔진을 받으므로, 어떤 엔진도 쓰지 않는 형식인 CSV는 애초에 선언할 자리가 없다. 각 엔진은 자기가 만드는 형식을 알고 있고, 시트/행/열 한도는 엔진이 아니라 그 형식에 속한다 — `XSSF`와 `SXSSF`가 한도를 공유하는 이유다.

그래서 CSV로 쓰는 것은 엔진이 아니라 시작 메서드로 고른다. `exportCsv()`/`exportSampleCsv()`는 POI 워크북을 만들지 않는 별도 경로를 타며, 그 경로가 형식을 `PxlFileFormat.CSV`로 고정하기 때문에 CSV 한도가 걸린다. 엔진이 정했을 것들은 — 스트리밍 윈도우를 포함해 — 그곳에서 아무 의미가 없다.

`PxlFileFormat`은 다운로드 응답을 조립할 때 쓰는 타입이다. 파일 확장자와 MIME content type을 갖고 있다.

```java
PxlFileFormat format = PxlExcelEngine.SXSSF.getFileFormat();                 // XLSX

response.setContentType(format.getContentType());                            // application/vnd.openxmlformats-...
response.setHeader("Content-Disposition",
        "attachment; filename=report." + format.getFilenameExtension());     // report.xlsx
```

이미 갖고 있는 것으로부터 두 값을 되찾을 수도 있다. 아래 조회는 전부 순수 조회로, **예외를 던지지 않고 `null`도 반환하지 않는다.**

| 조회                                         | 반환                                                                                        |
|--------------------------------------------|-------------------------------------------------------------------------------------------|
| `PxlFileFormat.fromPoiWorkbook(workbook)`  | 그 워크북이 담고 있는 형식. 스트리밍 리더의 워크북도 다른 OOXML과 똑같이 `XLSX`를 답한다. `CSV`는 대응하는 POI 워크북이 없어 반환되지 않는다 |
| `PxlExcelEngine.fromPoiWorkbook(workbook)` | 그 워크북 뒤의 writer. `XSSF`와 `SXSSF`를 구분한다                                                    |
| `PxlExcelEngine.fromWorkbookObject(Class)` | 클래스가 `@PxlWorkbook`으로 선언한 엔진                                                              |

---

## 지원 변수 타입

| 분류     | 타입                                                                                           |
|--------|----------------------------------------------------------------------------------------------|
| 기본형/래퍼 | `byte` `short` `int` `long` `float` `double` `char` `boolean` 및 각 래퍼 클래스, `String`           |
| 수치     | `BigInteger` `BigDecimal`                                                                    |
| 날짜·시각  | `Date` `LocalDate` `LocalTime` `LocalDateTime` `ZonedDateTime` `OffsetTime` `OffsetDateTime` |
| 기타     | `Enum`, `UUID`, 사용자 정의 클래스, 위 타입들의 `Collection`                                              |
| 실험적    | `Duration` `Period`                                                                          |

미지원 변수 타입 필드는 컬럼 메타를 해석하는 시점에 `PxlArgumentException`으로 실패한다(사용자 정의 클래스라면 `@PxlExportConverter`·`@PxlImportConverter` 또는 `String` 생성자를 갖춰야 지원 타입이 된다). 지원 타입이지만 셀 값을 그 타입으로 변환할 수 없을 때가 `PxlCellCodecException`이다.

### 타입별 동작 요약 — Export

> 아래 표는 셀에 타입이 있는 엑셀 export 기준이다. CSV에는 셀 타입이 없으므로 필드에 들어가는 것은 언제나 codec이 계산한 문자열이다 —
> `pattern` 없는 수치·날짜는 표시 형식을 얹은 serial이 아니라 그 codec의 문자열(`LocalDate`면 `2023-06-15`)로 기록되고, 아래의 quote-prefix 열은 CSV에 대응물이 없다.

| 타입                                            | Export 동작 / 특이사항·제한                                                                                                                                                                                              |
|------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `byte`·`short`·`int` + 래퍼 클래스               | 숫자 셀로 기록(표현 범위가 2^53 미만이라 안전)                                                                                                                                                                                    |
| `long`·`Long`                                  | 패턴 없으면 숫자 셀(double) → 2^53 초과 정밀도 손실.<br/>보존하려면 `pattern` 또는 `BigInteger`/`BigDecimal` 사용                                                                                                                        |
| `float`·`double` + 래퍼 클래스                   | 숫자 셀                                                                                                                                                                                                             |
| `char`·`Character`                             | 단일 문자 문자열로 기록. `exportTrim`은 적용되지 않음                                                                                                                                                                             |
| `boolean`·`Boolean`                            | `exportTrueString`/`exportFalseString`에 따른 문자열로 기록                                                                                                                                                               |
| `String`                                       | 텍스트 기록. `exportTrim`, 마스킹(`exportMasking`), `exportStringAsFormula`(선두 `=` → 수식), `exportStringAsPicture`(이미지) 옵션                                                                                                |
| `BigInteger`·`BigDecimal`                      | 항상 문자열 셀로 정밀도 보존 → Excel 정렬·수식·필터 대상에서 제외될 수 있음.<br/>`pattern` 지정 시 `DecimalFormat`으로 반올림 가능                                                                                                                     |
| `Date`·`LocalDate`·`LocalTime`·`LocalDateTime` | `pattern`/`exportPattern`·마스킹이 없으면 날짜 서식이 적용된 숫자 셀로 기록.<br/>적용 표시형식은 POI 내장 형식 코드(로케일 무관): `Date`·날짜=`m/d/yy`, 시각=`h:mm:ss`, 일시=`m/d/yy h:mm`.<br/>`pattern`/`exportPattern` 또는 마스킹을 지정하면 그 값을 살리기 위해 문자열 셀로 기록. |
| `ZonedDateTime`·`OffsetDateTime`·`OffsetTime`  | 위와 동일하게 패턴·마스킹이 없으면 숫자 셀로 기록.<br/>zone/offset을 보존하려면 offset 포함 `pattern`(예 `"yyyy-MM-dd HH:mm XXX"`)으로 문자열 출력                                |
| `Enum`                                         | `toString()` 오버라이드값 또는 상수명.<br/>`@PxlExportConverter`로 커스텀                                                                                                                                                       |
| 사용자 정의 클래스                                | `toString()` 재정의 또는 `@PxlExportConverter` 필요                        |
| `Collection`                                   | 구분자로 조인, `null` 요소 → 빈칸(예 `["a", null, "b"]`→`"a;;b"`).<br/>`exportStringAsPicture` 가능                                                                                                                        |
| `Duration`·`Period` (실험적)                    | 패턴 없으면 ISO-8601(`toString()`).<br/>패턴이면 `DurationFormatUtils`; `Period`는 현재 시각 기준 환산이라 근사값                                                                                                                       |
| `UUID`                                         | 언제나 canonical 소문자 형식(`toString()`)의 문자열 셀로 기록된다. 따라서 대문자 값도 소문자로 기록되어 되돌아오며, 값 자체는 동일하다.<br/>`exportTrim`·마스킹(`exportMasking`)은 그 문자열에 적용되고, `pattern`/`exportPattern`은 UUID에 의미가 없어 무시된다                                                       |

**Export 공통**

- `null` 값 및 빈/공백 `String` 값은 컬럼의 `exportNullString`으로 기록되며, 그 기본값은 빈 문자열 `""`이다.  
  즉 타입과 무관하게 `null` 필드는 기본적으로 빈 문자열이 든 문자열 셀로 export된다.  
  `exportNullString`으로 다른 문자열을 지정할 수 있다.
- Export는 기본 XLSX로 생성되며, `exportExcelEngine`로 `HSSF` 엔진(XLS)·`SXSSF` 엔진(스트리밍 XLSX)도 선택할 수 있다. CSV export는 미지원이며, 엔진은 엑셀 writer이므로 CSV를 지정할 자리 자체가 없다.
- 시트/컬럼 순서는 필드 선언 순서를 보장하지 않으므로, 순서가 중요하면 `exportOrder`를 지정한다.

### 타입별 동작 요약 — Import

| 타입                                            | Import 동작 / 특이사항·제한                                                                                                                                                                     |
|------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `byte`·`short`·`int`·`long` + 래퍼 클래스        | 숫자 셀/문자열 셀: `pattern` 지정 시 `DecimalFormat`로 파싱.<br/>타입 범위 초과 시 `PxlCellCodecException` 발생.<br/>소수부는 절단(예 `12.9`→`12`).<br/>`long`/`Long`은 숫자 셀이 double이라 2^53 초과 정밀도 손실 가능(문자열 셀이면 정확)  |
| `float`·`double` + 래퍼 클래스                   | 숫자 셀/문자열 셀: IEEE-754 정밀도 파싱.<br/>`NaN`·`Infinity`, 그리고 `float` 표현 범위를 넘는 값은 `PxlCellCodecException` 발생.<br/>단, `float` 언더플로(예 `1e-300`→`0.0f`)는 IEEE-754 정밀도 손실로 허용                     |
| `char`·`Character`                             | 문자열 셀: 첫 글자.<br/>숫자 셀: 정수/실수를 그대로 문자열화(`12`→`"12"`, `-3`→`"-3"`)한 뒤 첫 글자만 취한다. (`12`→`'1'`, `-3`→`'-'`).<br/>불리언 셀: `'1'`/`'0'`.<br/>빈 셀·빈 값: `Character`=`null`, `char`=`'\0'`              |
| `boolean`·`Boolean`                            | 문자열 셀: 먼저 `importTrueString`/`importFalseString`과 대소문자 무시 비교, 다음으로 내장 토큰 `true/false`·`t/f`·`y/n`·`yes/no`·`on/off`·`1/0`(대소문자 무시). 모두 불일치 → 예외(무음 `false` 아님).<br/>숫자 셀: 0이 아니면 `true` |
| `String`                                       | 문자열 셀: 값을 문자열로.<br/>숫자/불리언 셀: 문자열로 변환                                                                                                                                                   |
| `BigInteger`·`BigDecimal`                      | 문자열 셀: `new BigInteger/BigDecimal`로 정확 복원.<br/>숫자 셀: double(2^53) 정밀도 제한                                                                                                                |
| `Date`·`LocalDate`·`LocalTime`·`LocalDateTime` | 문자열 셀: `pattern`/`importPattern` 또는 고정 ISO-8601 기본 패턴으로 파싱.<br/>숫자 셀: 엑셀 날짜/시각 인식.<br/>불리언 셀: 지원하지 않아 예외 발생.                                                                            |
| `ZonedDateTime`·`OffsetDateTime`·`OffsetTime`  | 문자열 셀: `pattern` 또는 zone/offset 포함 ISO-8601로 파싱하며 zone/offset을 보존.<br/>offset/zone이 없는 문자열은 `pattern` 없이는 예외 발생.<br/>숫자 셀: 엑셀 날짜/시각을 시스템 기본존/오프셋 기준으로 읽는다.<br/>불리언 셀: 예외 발생.            |
| `Enum`                                         | `toString()` 오버라이드값 또는 상수명과 매칭(대소문자·공백 무시).<br/>`@PxlImportConverter`로 커스텀 변환                                                                                                           |
| 사용자 정의 클래스                                | `String` 단일인자 생성자 또는 `@PxlImportConverter` 필요                                                                                                                                           |
| `Collection`                                   | 구분자로 분리, 빈/null 요소의 위치 보존(예 `"a;;b"`→`["a", null, "b"]`).<br/>요소는 구체 클래스여야 하며 중첩 제네릭(`List<List<..>>`)·와일드카드(`List<? extends X>`)·raw type은 `PxlReflectionException` 발생.                |
| `Duration`·`Period` (실험적)                    | 문자열 셀: `pattern` 또는 ISO-8601.<br/>숫자 셀: 단위 고정(`Duration`=초, `Period`=일), 소수부 절단, 범위 초과 시 예외 발생.                                                                                         |
| `UUID`                                         | 문자열 셀: canonical 8-4-4-4-12 16진 형식만 허용하며 대소문자는 가리지 않는다(`123E4567-…`는 `123e4567-…`와 같은 값으로 바인딩).<br/>숫자 셀: 먼저 문자열로 렌더링하므로 "지원하지 않는 셀 타입"이 아니라 "유효하지 않은 값"으로 보고된다.<br/>불리언 셀: 예외 발생.<br/>`pattern`/`importPattern`은 UUID에 의미가 없어 무시된다. |

**Import 공통**

- 빈 셀 / 빈 값(Excel BLANK·없는 셀, CSV 빈 값)은 필드에 값을 설정하지 않는다 — 참조형은 `null`, primitive는 DTO 기본값(`0`/`false` 등)으로 설정.
- `pattern`/`importPattern`은 셀 값 **전체**와 맞아야 한다. 패턴이 앞부분만 읽는 값 — `"#,##0"`에 `"123abc"`·`"1e3"`, `"yyyy-MM-dd"`에 `"2024-01-02 xxx"` — 은 읽어낸 앞부분을 바인딩하지 않고 `PxlCellCodecException`을 낸다. 따라서 패턴을 지정한다고 해서 지정하지 않았을 때보다 더 관대해지는 일은 없다. (검사하는 것은 소비 여부뿐이다 — `"1,2,3"`처럼 그룹 구분자의 **위치**가 어긋난 값은 여전히 통과한다. `DecimalFormat`이 파싱 시 그룹 크기를 검증하지 않기 때문이다.)
- `importTrim`은 기본 `true`. `false`면 `String`만 공백을 보존하고, 다른 타입은 공백이 섞이면 파싱 실패/잘못된 값이 될 수 있다. `pattern`이 있으면 뒤쪽 공백(`"123 "`)도 소비되지 않은 입력이므로 거부된다.
- Streaming Reader(`importUsingStreamReader`, XLSX 전용)는 수식 셀을 평가하지 못하며 헤더 행 위치를 정확히 지정해야 한다.

---

## 애노테이션

애노테이션이 붙은 필드/클래스만 바인딩된다.
옵션 객체를 `.override()` 단계로 넘겨 런타임 값을 주면 그 값이 애노테이션 값을 오버라이드한다.

### `@PxlWorkbook` (클래스 대상)

| 속성                                                                | 기본값        | 설명                                                                                                                                                                                                                                |
|-------------------------------------------------------------------|------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `exportExcelEngine`                                               | `XSSF`     | 워크북을 쓰는 POI 엔진(`PxlExcelEngine`): `XSSF`=XLSX(기본), `HSSF`=XLS, `SXSSF`=스트리밍 XLSX.<br/>형식이 아니라 writer를 고르는 속성이라 `XSSF`와 `SXSSF`는 똑같이 `.xlsx`를 만든다. CSV는 엔진이 아니므로 여기에 지정할 수 없다. |
| `exportPassword`                                                  | `""`       | Export 시 설정할 문서보호 비밀번호.<br/>`toFile(...)`·`toStream(...)`에만 적용되고 `toWorkbook()`에는 적용되지 않는다.<br/>CSV export는 평문을 쓰는 대신 **거부**한다(`PxlArgumentException`). |
| `exportDataValidation`                                            | `true`     | 쓰려는 객체에 대해 Bean Validation을 수행할지 여부.<br/>엑셀의 "데이터 유효성" 기능과는 무관하다 — 위 `importDataValidation` 참조 |
| `exportSXSSFRowAccessWindowSize`                                  | `100`      | SXSSF Export 시 rowAccessWindowSize                                                                                                                                                                                                 |
| `exportCsvCharset`                                                | `""`→`"UTF-8"` | CSV 전용. Export할 CSV의 문자 인코딩(워크북의 모든 시트 공통). 개별 시트는 `@PxlSheet(exportCsvCharset)`으로 따로 설정할 수 있다                                                                                                        |
| `exportCsvDelimiter`                                              | `'\0'`→`','` | CSV 전용. Export할 CSV의 구분자(워크북의 모든 시트 공통, `char`). 개별 시트는 `@PxlSheet(exportCsvDelimiter)`으로 따로 설정할 수 있다                                                                                              |
| `exportCsvBom`                                                    | `UNSPECIFIED`→`false` | CSV 전용. 출력 앞에 BOM을 붙일지 여부(`PxlOptionalBoolean`). 개별 시트는 `@PxlSheet(exportCsvBom)`으로 따로 설정할 수 있다.<br/>UTF-8/UTF-16LE/UTF-16BE에서만 적용되고, 그 외 문자셋에서는 조용히 생략된다(↓ *제한사항*) |
| `exportWorkbookRequiredHeaderCellStyler`                          | (미지정)      | 필수 헤더 셀 스타일 (미지정/적용불가 시 `PxlHeaderRequiredStyler`)                                                                                                                                                                  |
| `exportWorkbookOptionalHeaderCellStyler`                          | (미지정)      | 선택 헤더 셀 스타일 (미지정/적용불가 시 `PxlHeaderOptionalStyler`)                                                                                                                                                                  |
| `exportWorkbookDataCellStyler`                                    | (미지정)      | 데이터 셀 스타일 (미지정/적용불가 시 `PxlDataVerticalCenterTextStyler`)                                                                                                                                                             |
| `exportI18nBaseName` / `exportI18nLanguage` / `exportI18nCountry` | `""`/`"en"`/`""` | Export 시 다국어 ResourceBundle의 base name / language / country                                                                                                                                                                    |
| `importPassword`                                                  | `""`       | Import 시 문서보호를 해제할 비밀번호                                                                                                                                                                                                |
| `importDataValidation`                                            | `true`     | Import 결과에 대해 Bean Validation(`@NotNull`·`@Size` 등 — ↓ *유효성 검사*)을 수행할지 여부.<br/>엑셀의 "데이터 유효성" 기능과는 무관하다 — 그쪽은 `@PxlColumn(exportOptionItems)`이 만드는 드롭다운이다(↓ *컬럼: 드롭다운 Export*) |
| `importUsingStreamReader`                                         | `false`    | Import 시 Streaming Reader 사용 여부 (XLSX 전용)                                                                                                                                                                                    |
| `importStreamReaderRowCacheSize`                                  | `100`      | Streaming Reader의 row cache size                                                                                                                                                                                                   |
| `importStreamReaderBufferSize`                                    | `4096`     | Streaming Reader의 buffer size                                                                                                                                                                                                      |
| `importCsvCharset`                                                | `""`→`"UTF-8"` | CSV 전용. Import할 CSV의 문자 인코딩(워크북의 모든 시트 공통). 개별 시트는 `@PxlSheet(importCsvCharset)`으로 따로 설정할 수 있다.<br/>선두 BOM 자동 처리(UTF-8/UTF-16LE/BE의 BOM 제거, `UTF-16`(auto)의 BOM은 엔디안 판별에 사용)             |
| `importCsvDelimiter`                                              | `'\0'`→`','` | CSV 전용. Import할 CSV의 구분자(워크북의 모든 시트 공통, `char`). 개별 시트는 `@PxlSheet(importCsvDelimiter)`으로 따로 설정할 수 있다                                                                                                              |
| `importI18nBaseName` / `importI18nLanguage` / `importI18nCountry` | `""`/`"en"`/`""` | Import 시 다국어 ResourceBundle의 base name / language / country                                                                                                                                                                    |

### `@PxlWorkbookName` (필드 대상)

워크북 이름을 담을 `String` 필드에 붙이며 불필요하면 생략 가능하다.  
`String` 타입이 아닌 필드에 붙이면 `PxlDataException` 발생한다.  
import 전용: 워크북 형식 import 시, 빌더의 `.workbookName(...)`으로 넘긴 이름이 이 필드에 채워진다.
이름을 주지 않으면 파일에서 읽는 경우 그 파일의 확장자를 제거한 파일명으로 설정된다.
export(및 샘플 export)에서는 사용되지 않는다.

### `@PxlSheet` (필드 대상)

`Collection` 타입 필드에 붙여 시트로 바인딩한다.  
인덱스 속성의 기본값 `0`은 "auto"(첫 행/열 자동)를 뜻한다.

| 속성                                                                                                          | 기본값     | 설명                                                                                                             |
|-------------------------------------------------------------------------------------------------------------|---------|------------------------------------------------------------------------------------------------------------------|
| `name`                                                                                                      | 필드명     | 시트 이름(배열). 실제 시트명과 일치해야 바인딩됨(공백·대소문자 무시).<br/>배열로 지정 시 그중 하나만 존재해야 함 |
| `exportEnabled` / `exportSampleEnabled`                                                                     | `true`  | Export / 샘플 Export 사용 여부                                                                                   |
| `exportOverrideSuperClassSheet`                                                                             | `false` | 슈퍼클래스의 동일 시트명 필드를 override할지(대소문자 무시)                                                      |
| `exportRowHeightInPoints`                                                                                   | `-1.0`  | 시트 내 행 높이(point). 미설정 시 기본 높이                                                                      |
| `exportOrder`                                                                                               | `""`    | 시트 생성 순서 키 (문자열 비교, 아래 *Export 순서* 참고)                                                         |
| `exportGroupingFieldName`                                                                                   | `""`    | 이 필드 값으로 그룹핑하여 여러 시트로 분할                                                                       |
| `exportHeaderRowIndex` / `exportFirstDataRowIndex` / `exportLastDataRowIndex`                               | `0`     | Export 시 Header/시작/끝 데이터 행 (1-based)                                                                     |
| `exportFirstDataColumnIndex` / `exportLastDataColumnIndex`                                                  | `0`     | Export 시 시작/끝 데이터 열 (1-based)                                                                            |
| `exportIfNull`                                                                                              | `false` | 필드가 null일 때 시트 생성 여부                                                                                  |
| `exportIfEmpty`                                                                                             | `true`  | 필드가 비었을 때 시트 생성 여부                                                                                  |
| `exportColumnFilter`                                                                                        | `false` | 필터 적용 여부                                                                                                   |
| `exportCsvCharset`                                                                                          | (상속)    | CSV 전용. 이 시트의 CSV를 쓸 문자 인코딩. 빈 값(`""`)이면 워크북 값을 상속                                        |
| `exportCsvDelimiter`                                                                                        | (상속)    | CSV 전용. 이 시트의 CSV를 쓸 구분자(`char`). NUL(`'\0'`)이면 워크북 값을 상속                                     |
| `exportCsvBom`                                                                                              | (상속)    | CSV 전용. 이 시트의 CSV 앞에 BOM을 붙일지(`PxlOptionalBoolean`). `UNSPECIFIED`면 워크북 값을 상속하고, `FALSE`는 워크북이 켠 BOM을 끈다 — 평범한 `boolean`으로는 표현할 수 없는 경우다 |
| `exportSheetRequiredHeaderCellStyler` / `exportSheetOptionalHeaderCellStyler` / `exportSheetDataCellStyler` | (미지정)   | 시트 단위 셀 스타일 (미지정 시 Workbook 단위로 위임)                                                             |
| `importEnabled`                                                                                             | `true`  | Import 사용 여부                                                                                                 |
| `importOverrideSuperClassSheet`                                                                             | `false` | 슈퍼클래스의 동일 시트명 필드를 override할지 여부(대소문자 무시)                                                 |
| `importExcludeHiddenRows` / `importExcludeHiddenColumns`                                                    | `false` | 숨겨진 행/열 제외 여부                                                                                           |
| `importEachCellOfMergedRegion`                                                                              | `false` | 병합 셀을 개별 셀에 동일 값으로 처리할지                                                                         |
| `importHeaderRowIndex` / `importFirstDataRowIndex` / `importLastDataRowIndex`                               | `0`     | Import 시 Header/시작/끝 데이터 행 (1-based, 아래 *인덱스 규칙* 참고)                                            |
| `importFirstDataColumnIndex` / `importLastDataColumnIndex`                                                  | `0`     | Import 시 시작/끝 데이터 열 (1-based)                                                                            |
| `importCsvCharset`                                                                                          | (상속)    | CSV 전용. 이 시트의 CSV를 읽을 문자 인코딩. 빈 값(`""`)이면 워크북 값을 상속                                     |
| `importCsvDelimiter`                                                                                        | (상속)    | CSV 전용. 이 시트의 CSV를 읽을 구분자(`char`). NUL(`'\0'`)이면 워크북 값을 상속                                  |

### `@PxlRowIndex` (필드 대상)

행 인덱스를 담을 필드에 붙이며 불필요하면 생략 가능하다.  
필드 타입은 `byte`·`short`·`int`·`long` 및 각 래퍼 클래스(`Byte`·`Short`·`Integer`·`Long`)를 지원하며, 그 외 타입이면 `PxlArgumentException`으로 실패한다.  
채워지는 값은 가져온 행의 1-based 스프레드시트 행 번호(스프레드시트 UI에 표시되는 행 번호이자 `importHeaderRowIndex` 등 인덱스 속성과 같은 기준)다.  
기본 구성(헤더가 첫 행)에서는 데이터 행이 2·3·4…이며, `importHeaderRowIndex`로 헤더를 아래로 옮기거나 헤더 위에 제목 행이 있으면 그만큼 커진 절대 행 번호가 들어간다.  
import 전용: export(및 샘플 export)에서는 사용되지 않는다.

### `@PxlColumn` (필드 대상)

| 속성                                                                                                             | 기본값        | 설명                                                                                                             |
|----------------------------------------------------------------------------------------------------------------|------------|----------------------------------------------------------------------------------------------------------------|
| `name`                                                                                                         | 필드명        | 열 이름(배열). 실제 열명과 일치해야 바인딩됨(공백은 무시, 대소문자는 구분).<br/>배열로 지정 시 그중 하나만 존재해야 함                                                         |
| `pattern`                                                                                                      | `""`       | `exportPattern` / `importPattern`이 비었을 때 쓰는 공통 폴백 패턴.<br/>파싱에 쓰일 때는 값 전체와 맞아야 한다(↑ *Import 공통*) |
| `collectionSeparator`                                                                                          | `";"`      | `exportCollectionSeparator` / `importCollectionSeparator`가 비었을 때 쓰는 공통 폴백                                      |
| `exportEnabled` / `exportSampleEnabled`                                                                        | `true`     | Export / 샘플 Export 사용 여부                                                                                       |
| `exportSample`                                                                                                 | `""`       | 샘플 Export 시 셀에 넣을 값.<br/>컬럼 타입으로 파싱되므로 그 타입이 받아들이는 값이어야 함(아니면 `PxlCellCodecException`).<br/>지정하지 않으면 샘플 셀에 `exportNullString`이 들어감 |
| `exportTrim`                                                                                                   | `false`    | Export 시 문자열 trim 여부 (`char`/`Character` 컬럼에는 미적용)                                                             |
| `exportPattern`                                                                                                | `""`       | Export 형식(수치=`DecimalFormat`, 날짜·시각=`DateTimeFormat`, `Duration`/`Period`=`DurationFormatUtils`). |
| `exportColumnWidth`                                                                                            | `0`(auto)  | 열 너비. 기본값(0=`autoSizeColumn`)은 전체 행을 실측하므로 **대량 데이터 export 시 성능 저하**(↓ *제약* 절).<br/>행이 많으면 고정 너비 지정 권장             |
| `exportCollectionSeparator`                                                                                    | `""`       | Collection 요소 구분 문자열.                                                                 |
| `exportOverrideSuperClassColumn`                                                                               | `false`    | 슈퍼클래스의 동일 컬럼명 필드를 override할지                                                                                   |
| `exportOrder`                                                                                                  | `""`       | 컬럼 생성 순서 키 (문자열 비교, 아래 *Export 순서* 참고)                                                                         |
| `exportMasking`                                                                                                | `""`       | 마스킹할 부분의 정규식.<br/>`char`/`Character` 컬럼에는 미적용                                                             |
| `exportOptionItems`                                                                                            | `{}`       | 선택 가능한 옵션 목록(드롭다운)                                                                                             |
| `exportEnumDropDownListStyle`                                                                                  | `SET`      | Enum 필드를 드롭다운으로 설정할 스타일 (`SET` / `SORTED_SET` / `NONE`)                                                        |
| `exportNullString`                                                                                             | `""`       | null 값(및 빈/공백 `String`)을 export할 때 쓸 문자열. 기본은 빈 문자열이 든 문자열 셀(blank 셀 아님)                                       |
| `exportTrueString` / `exportFalseString`                                                                       | `"true"`/`"false"` | 참/거짓을 export할 때 쓸 문자열.<br/>커스텀 값을 다시 import하려면 `importTrueString`/`importFalseString`도 같은 값으로 지정               |
| `exportStringAsPicture`                                                                                        | `false`    | 이미지 URL 문자열을 셀에 이미지로 삽입 |
| `exportStringAsFormula`                                                                                        | `false`    | 수식 문자열(선두 `=`)을 계산하여 셀에 적용.<br/>`exportStringAsPicture`와 함께 지정하면 이쪽이 우선 |
| `exportColumnRequiredHeaderCellStyler` / `exportColumnOptionalHeaderCellStyler` / `exportColumnDataCellStyler` | (미지정)      | 컬럼 단위 셀 스타일 (미지정 시 Sheet 단위로 위임)                                                                               |
| `importEnabled`                                                                                                | `true`     | Import 사용 여부                                                                                                   |
| `importTrim`                                                                                                   | `true`     | Import 시 문자열 trim 여부.<br/>단, `false`이면 숫자·날짜·`Boolean` 등은 공백 탓에 파싱 실패/오값이 될 수 있다                               |
| `importUnique`                                                                                                 | `false`    | Import 시 열 값들의 유일성 검사 여부                                                                                       |
| `importPattern`                                                                                                | `""`       | Import 형식(수치=`DecimalFormat`, 날짜·시각=`DateTimeFormat`, `Duration`/`Period`=`DurationFormatUtils`).<br/>셀 값 전체와 맞아야 하며, 남는 문자가 있으면 그 값은 무효다.<br/>날짜·시각은 실패 시 기본 패턴 폴백(부분 일치도 실패로 쳐서 폴백 포맷터로 넘어간다), `Duration`/`Period`도 같은 방식으로 ISO-8601로 폴백한다 |
| `importTrueString` / `importFalseString`                                                                       | `"true"`/`"false"` | `String` 열: 불리언 셀을 이 문자열로 렌더링.<br/>`Boolean` 열: 이 문자열(대소문자 무시)을 참/거짓으로 해석(내장 토큰보다 우선)                              |
| `importCollectionSeparator`                                                                                    | `""`       | Cell 값을 Collection 요소로 분리할 구분자.<br/>리터럴 전체 문자열(`"::"`·`", "` 등 다중문자 가능).              |
| `importOverrideSuperClassColumn`                                                                               | `false`    | 슈퍼클래스의 동일 컬럼명 필드를 override할지                                                                                   |

### `@PxlExportConverter` / `@PxlImportConverter` (메서드 대상)

Enum 또는 사용자 정의 클래스의 커스텀 객체 ↔ String 변환을 지정한다.

```java
// 값 → 문자열
@PxlExportConverter
public String pxlExportConverter() {
    return ...;
}

// 문자열 → 값 (static, 반환 타입은 대상 타입)
@PxlImportConverter
public static EnumOrObjectType pxlImportConverter(final String str) {
    return ...;
}
```

> **제약:**   
> `@PxlExportConverter`는 인스턴스 메서드(`String pxlExportConverter()`) 또는 대상 값을 인자로 받는
> `static` 메서드(`static String pxlExportConverter(Type value)`) 둘 다 가능하며, 반환 타입은 `String`이어야 한다.  
> `@PxlImportConverter` 메서드는 반드시 `static` 이어야 하고, 반환 타입은 대상 타입(enum/사용자 정의 클래스)과 일치해야 한다(문자열로부터 새 값을
> 생성하는 팩토리이므로 인스턴스 메서드는 지원하지 않는다).

---

## 옵션 오버라이드

### 옵션 사용법

옵션 객체(`PxlExportWorkbookOption`·`PxlImportWorkbookOption`)는 `@PxlWorkbook`·`@PxlSheet`·`@PxlColumn` 애노테이션으로 선언한 값을 런타임에 오버라이드할 값들의 묶음으로, 워크북 단위로 만든다.  
이 객체를 마지막(실행) 단계 전 `.override(...)` 구성 단계에 넘기면 담긴 값이 애노테이션 값을 덮어쓴다 — 지정하지 않은 필드는 애노테이션 값(없으면 기본값)을 그대로 따른다.  
`.override(...)` 단계 자체도 생략할 수 있으며, 생략하면 애노테이션 값을 그대로 쓴다.

```java
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;

// 1) 옵션 객체를 빌더로 만든다 (워크북 단위)
PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
        .exportPassword("secret")          // 문서 암호화
        .exportDataValidation(false)       // export 데이터 유효성 검사 끄기
        .build();

// 2) 실행 전 .override(...) 으로 넘겨 애노테이션 값을 오버라이드한다
pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .override(exportOption)                                   // @PxlWorkbook(exportPassword=...) 등을 오버라이드
   .toFile(new File("secured.xlsx"));
```

```java
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

// 1) 옵션 객체를 빌더로 만든다 (워크북 단위)
PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
        .importDataValidation(false)       // import 결과 유효성 검사 끄기 (기본 true)
        .importUsingStreamReader(true)     // 대용량 XLSX 스트리밍
        .build();

// 2) 실행 전 .override(...) 으로 넘겨 애노테이션 값을 오버라이드한다
List<Employee> rows = pxl.importExcel()
                         .override(importOption)             // 애노테이션 값을 오버라이드
                         .sheet(Employee.class, "Employees")
                         .fromFile(file);
```

### 옵션 구조

옵션은 워크북 한 단계가 아니라 워크북 → 시트 → 컬럼의 3단계 트리로 구성해 한 번의 `.override(...)`로 넘길 수 있다.

- 워크북 옵션의 `exportSheetOptions`/`importSheetOptions`에 시트 옵션(`Pxl{Export,Import}SheetOption`)을 담는다.
- 시트 옵션의 `exportColumnOptions`/`importColumnOptions`에 컬럼 옵션(`Pxl{Export,Import}ColumnOption`)을 담는다.
- 자식 옵션은 빌더의 리스트 세터(`.importColumnOptions(List)` 등)나 `add*Option(...)` 메서드로 붙인다.
- 매칭 키: 시트 옵션은 `fieldName`(워크북 클래스의 `@PxlSheet` 필드명), 컬럼 옵션은 `fieldName`(행 클래스의 `@PxlColumn` 필드명)으로 대상에 연결된다.  
  시트 옵션의 `fieldName`을 생략하면 와일드카드(`*`)로 모든 시트에 적용되며, 단일 시트 폼(`sheet(...)`)에서는 이 방식을 쓴다.
- 런타임 이름 변경: 시트 옵션은 `exportSheetNames`/`importSheetNames`로 `@PxlSheet(name)`을, 컬럼 옵션은 `exportColumnNames`/`importColumnNames`로 `@PxlColumn(name)`을 오버라이드한다(애노테이션과 마찬가지로 리스트라 별칭도 그대로 쓸 수 있다).  
  애노테이션의 `name`을 대체하는 값이므로 이들 역시 번들 키다 — i18n을 켜면 오버라이드한 이름도 매칭·출력 전에 번역된다(↓ [i18n](#i18n)).
- 지정하지 않은 레벨·필드는 애노테이션 값(없으면 기본값)을 그대로 따른다.

```java
// 워크북 → 시트(와일드카드) → 컬럼(age) 3단 오버라이드: age 컬럼만 바인딩에서 제외
PxlImportColumnOption ageColumn = PxlImportColumnOption.builder()
        .fieldName("age")                                   // 컬럼(@PxlColumn 필드명)으로 매칭
        .importEnabled(false)                               // 이 컬럼은 바인딩하지 않음
        .build();
PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()   // fieldName 생략 → 와일드카드(모든 시트)
        .importColumnOptions(Arrays.asList(ageColumn))
        .build();
PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
        .importSheetOptions(Arrays.asList(sheetOption))
        .build();

List<Employee> rows = pxl.importExcel()
                         .override(option)                          // 트리 전체를 한 번에 넘긴다
                         .sheet(Employee.class, "People")
                         .fromFile(file);
// → age는 바인딩되지 않아 초기값 그대로(Integer이므로 null), 나머지 컬럼은 정상 바인딩
```

export도 `exportSheetOptions`/`exportColumnOptions`로 동일하게 구성한다.

---

## 유효성 검사

### 표준 제약

필드에 유효성 검사 애노테이션(`@NotNull`, `@NotEmpty`, `@NotBlank`, `@Valid`, ...)을 붙이면 import 결과 객체에 대해 유효성 검사하고,
실패 시 `PxlValidationException`을 발생시킨다.

```java
@NotBlank
@PxlColumn(name = "Name")
private String name;

// 컬렉션 자체에 건 제약. 행 검사는 @Valid 유무와 무관하게 수행된다
@NotEmpty
@PxlSheet(name = "Employees")
private List<Employee> employees;
```

행 객체는 PXL이 항상 직접 검사한다 — 행 단위로 검사하며 예외에 시트명(import은 행 번호까지)을 붙인다. 따라서 `@PxlSheet` 필드에
`@Valid`를 붙이지 않아도 행은 검사된다. 반대로 PXL이 처리하지 않는 시트 — export의 `exportEnabled = false`, import의
`importEnabled = false` — 의 행은 `@Valid`를 붙여도 검사되지 않는다. 즉 행 검사는 실제로 쓰거나 읽는 시트에 대해서만 이뤄진다.

유효성 검사는 `@PxlWorkbook(importDataValidation = ...)`로 제어되며(기본 `true`), 끄면 유효성 검사가 수행되지 않는다.
단, 이 속성은 필드의 값이 제약을 만족하는지에만 관여한다.  
`@NotNull`·`@NotEmpty`·`@NotBlank`로 필수 지정한 컬럼이 시트 헤더에 없으면 유효성 검사를 꺼도 예외가 발생한다 — 필수 컬럼이 실제로 존재하는지는 값 검증과 별개로 항상 확인하기 때문이다.  
Bean Validation 구현체/EL이 클래스패스에 없으면 유효성 검사는 조용히 비활성화된다([구성](#구성) 참고).

### 커스텀 제약 — `@PxlByteSize`

표준 유효성 검사 제약(`@NotNull`·`@Size` 등)과 함께, PXL은 문자열의 바이트 길이를 유효성 검사하는 `@PxlByteSize`를 제공한다.  
`@Size`가 문자 수를 세는 것과 달리, `@PxlByteSize`는 지정 charset으로 인코딩한 바이트 수를 유효성 검사한다.  
DB `VARCHAR(n BYTE)`처럼 한글 1글자가 여러 바이트를 차지하는 컬럼 제약을 그대로 표현할 때 유용하다.

```java
import io.github.hclimkr.pxl.constraint.PxlByteSize;

// UTF-8 기준 최대 30바이트 (한글 10자 = 30바이트)
@PxlByteSize(max = 30)                              // charset 기본값 "UTF-8"
@PxlColumn(name = "Name")
private String name;

// 범위 지정 + charset 명시 (EUC-KR: 한글 1자 = 2바이트)
@PxlByteSize(min = 4, max = 20, charset = "EUC-KR")
@PxlColumn(name = "Code")
private String code;
```

- 대상 타입은 `CharSequence`(주로 `String`). `null`은 유효로 간주한다.
- `min`/`max`는 바이트 길이(경계 포함). `charset` 미지정 시 `"UTF-8"`.
- `importDataValidation`이 켜져 있을 때 import 결과 객체에 대해 유효성 검사되고, 실패 시 `PxlValidationException` 발생.

---

## 예외

### 예외 종류

경계를 넘는 예외는 모두 `Pxl` 경계에서 checked `PxlException` 계열로 정규화된다.  
예외의 메시지 문구는 다국어로 지역화된다.  
기본은 영어이며 한국어를 제공하고 `Pxl.setMessageLocale(Locale)`로 언어를 전역 지정한다.  
자세한 내용은 [i18n](#i18n)의 "예외·진단 메시지 언어"를 참조한다.

| 예외                        | 발생 시점                                                                                                                              |
|---------------------------|----------------------------------------------------------------------------------------------------------------------------------------|
| `PxlException`            | abstract 베이스 타입 — 직접 던져지지 않으며, 모든 checked Pxl 예외의 공통 상위 타입                                                    |
| `PxlCellCodecException`   | 셀 값을 대상 타입으로 변환할 수 없을 때                                                                                                |
| `PxlValidationException`  | 유효성 검사 실패 시                                                                                                                    |
| `PxlReflectionException`  | 리플렉션 실패 시                                                                                                                       |
| `PxlArgumentException`    | 인자/애노테이션 설정 오류 시                                                                                                           |
| `PxlNullPointerException` | 필수(non-null) 인자가 `null`일 때                                                                                                      |
| `PxlDataException`        | 워크북/CSV 형태가 기대와 다를 때 — 대상 시트·컬럼을 못 찾거나 중복, 시트/행/열 수 한도 초과, 헤더 행·컬럼 없음, export할 데이터 없음, export할 컬렉션에 `null` 원소 포함, 이름·클래스·스트림 개수 불일치, 행/열 인덱스 설정 위반 |
| `PxlIOException`          | I/O 오류일 때 — 파일·스트림을 열거나 읽지 못함(파일 없음·미지원 포맷·`importPassword` 불일치로 복호화 실패 포함), 쓰기 실패             |
| `PxlI18nException`        | i18n ResourceBundle을 못 찾을 때                                                                                                       |
| `PxlSystemException`      | PXL이 분류하지 않은 실패일 때 — 위 타입에 해당하지 않는 예외(POI의 예기치 못한 런타임 실패 등)를 경계에서 감싼 것(원인은 `getCause()`) |
| `PxlRuntimeException`     | unchecked 예외 — 현재 미사용                                                                                                           |

`PxlNullPointerException`·`PxlValidationException` 등 모든 checked 예외는 `PxlException`의 하위 타입이다.  
`PxlException` 자체는 abstract이므로 실제로 잡히는 것은 항상 구체 하위 타입이다 — 분류된 실패는 대응 타입으로, 그 밖의 실패는 `PxlSystemException`으로 온다.  
따라서 마지막(실행) 단계 메서드의 `throws PxlException` 선언은 그대로 유효한 계약이며, 호출부는 `catch (PxlException e)`로 일괄 처리하거나 필요한 하위 타입만 골라 잡을 수 있다.  
단, `PxlRuntimeException`은 예외로 unchecked(`RuntimeException` 계열)이며 `PxlException`의 하위가 아니며 현재는 사용하지 않고 있다.  

### 예외·진단 메시지 언어

PXL이 던지는 예외 메시지와 진단 로그 문구는 다국어를 지원한다.  
이 문구들은 라이브러리가 아티팩트에 동봉한 번들(`pxl-messages`)에서 해석되며, 기본은 영어이고 한국어를 제공한다.

- 프로세스 전역 locale로 결정된다.  
  기본은 JVM 기본 locale(`Locale.getDefault()`), `Pxl.setMessageLocale(Locale)`로 전역 지정하고 `Pxl.resetMessageLocale()`로 해제한다.  
  매칭되는 번역이 없으면 영어로 폴백한다.
- 콘텐츠 i18n(시트/컬럼명, `@PxlWorkbook` 기반·워크북별)과 독립적이다. 영문 산출물을 만들면서 서버 로그의 예외는 한국어로 보거나 그 반대도 가능하다.

```java
Pxl.setMessageLocale(Locale.ENGLISH);   // 이후 예외/진단 문구를 영어로
// ...
Pxl.resetMessageLocale();               // JVM 기본 locale로 복귀
```

---

## i18n

`@PxlWorkbook`의 `export/importI18nBaseName`, `export/importI18nLanguage`, `export/importI18nCountry`로 `ResourceBundle`을
지정하면 시트명·컬럼명을 번역해 매칭/출력한다(UTF-8 properties 지원).  
`@PxlColumn`/`@PxlSheet`의 `name` 값이 번들의 키가 된다. 평범한 `ResourceBundle` 키이고 번들은 대개 애플리케이션의 다른
메시지와 공유하므로, 단순한 한 단어 대신 네임스페이스를 둔 키(`staff.column.role`)로 짓는다.

i18n은 기본적으로 비활성(opt-in) 이다.  
`export/importI18nBaseName`을 명시적으로 지정(또는 옵션에 `ResourceBundle` 주입)했을 때만 동작하며, base name이 비어 있으면 번들을 로드하지 않아 이름이 그대로 사용된다.  
base name을 지정했으나 해당 `ResourceBundle`을 찾지 못하면 `PxlI18nException`으로 실패한다.

base name을 지정하는 대신, 이미 갖고 있는 번들을 워크북 옵션의 `exportResourceBundle`/`importResourceBundle`로 직접 넘길 수도 있다.  
주입한 번들이 애노테이션보다 우선하며, 이 경우 `export/importI18nBaseName`·`Language`·`Country` 3종은 아예 로드되지 않으므로 `PxlI18nException`이 날 여지도 없다.  
애노테이션으로는 지목할 수 없는 곳에서 번들이 오는 경우 — 컨테이너가 관리하는 `MessageSource`, 요청마다 달라지는 locale 등 — 에 쓴다.

```java
ResourceBundle bundle = ResourceBundle.getBundle("messages", userLocale);   // 애플리케이션이 결정한 번들

List<Person> rows = pxl.importExcel()
                       .override(PxlImportWorkbookOption.builder()
                                                        .importResourceBundle(bundle)   // @PxlWorkbook(importI18nBaseName)보다 우선
                                                        .build())
                       .sheet(Person.class, "staff.sheet")
                       .fromFile(file);
```

`src/main/resources/messages.properties` — base 번들이며 UTF-8로 읽는다.

```properties
staff.sheet=Staff
staff.column.role=Role
staff.column.fullName=Full Name
```

`src/main/resources/messages_ko.properties` — `exportI18nLanguage = "ko"`일 때 대신 선택되는 변형.

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

    @PxlSheet(name = "staff.sheet")                 // 헤더/시트명이 "Staff"로 번역됨
    private List<Person> people;
}

public class Person {

    @PxlColumn(name = "staff.column.role")          // 헤더가 "Role"로 번역됨
    private String role;

    @PxlColumn(name = "staff.column.fullName")      // 헤더가 "Full Name"으로 번역됨
    private String fullName;
}
```

이름 외에 `@PxlColumn`의 두 속성도 export 시 같은 번들을 거친다.

- **`exportSample`** — `String`/enum 컬럼에서 번역된다. 그 타입의 `Collection` 컬럼이면 샘플이 `exportCollectionSeparator`로 나뉜 **원소별 키**이므로 **원소마다 따로 번역**한다. 구분자가 번들 값 안에 박히지 않고 애노테이션에 남는다. enum 샘플은 번역 후 다시 상수로 파싱되므로 셀에는 정규 이름이 들어간다.
- **`exportOptionItems`** — `String` 컬럼에서 번역되어, 드롭다운이 셀에 든 것과 같은 텍스트를 제시한다. 그 외 타입은 값을 정규 형태로 쓰므로 항목(및 항목 미지정 시 쓰이는 enum 상수)을 적힌 그대로 둔다. 번역하면 실제로 쓰인 값이 Excel이 검사하는 목록 밖으로 벗어난다.

`messages.properties`:

```properties
staff.column.role=Role
staff.column.roles=Roles
staff.role.admin=Administrator
staff.role.user=User
```

```java
@PxlColumn(name = "staff.column.role", exportSample = "staff.role.admin",
        exportOptionItems = {"staff.role.admin", "staff.role.user"})   // 샘플 셀 "Administrator", 드롭다운 Administrator/User
private String role;

@PxlColumn(name = "staff.column.roles", exportSample = "staff.role.admin;staff.role.user")   // 샘플 셀 "Administrator;User"
private List<String> roles;
```

> `i18nBaseName`을 지정했을 때만 동작한다(기본은 미번역). 지정했는데 번들을 못 찾으면 조용히 넘어가지 않고 `PxlI18nException` 발생.

---

## 시트 내의 행/열 인덱스 규칙

`@PxlSheet`의 `exportFirstDataColumnIndex`, `importHeaderRowIndex` 등 인덱스 속성은 모두 1-based이며,
기본값 `0`은 auto(첫/마지막 자동)를 뜻한다.

| 속성                     | 기본값     | 제약                                          |
|------------------------|---------|---------------------------------------------|
| `HeaderRowIndex`       | 첫 행     | `FirstDataRowIndex`보다 작아야 함                 |
| `FirstDataRowIndex`    | 헤더 다음 행 | `HeaderRowIndex`보다 크고 `LastDataRowIndex` 이하 |
| `LastDataRowIndex`     | 마지막 행   | `FirstDataRowIndex` 이상                      |
| `FirstDataColumnIndex` | 첫 열     | `LastDataColumnIndex` 이하                    |
| `LastDataColumnIndex`  | 마지막 열   | `FirstDataColumnIndex` 이상                   |

> Export 시 검증:  
> - `exportLastDataColumnIndex`로 지정한 열 범위가 export할 컬럼 수보다 작으면(일부 컬럼 누락) 예외가 발생한다.
> - 같은 시트 안에 중복된 컬럼명이 있으면 예외가 발생한다.

> CSV export도 다섯 속성을 엑셀과 똑같이 지킨다 — 좌표 해석이 공용 메타에서 이뤄지기 때문이다.
> 다만 텍스트 파일에서 좌표가 어떤 모양이 되는지는 알아 둘 만하다. 헤더보다 앞선 행이나 `FirstDataColumnIndex` 이전의 열은
> 빈 필드로 기록된다 — 아무것도 없는 빈 줄이 아니라 빈 필드로 이뤄진 레코드(`"",,,`)다. 의도된 동작이다.
> PXL의 CSV import는 빈 줄을 무시하므로, 빈 줄로 쓰면 다시 읽을 때 헤더가 위로 당겨져 왕복이 깨진다.
  
> Import 시 빈 데이터 범위 동작:  
> - 데이터 행 범위가 실제 데이터와 겹치지 않거나 비게 되는 설정(예: `importFirstDataRowIndex`를 실제 데이터보다 큰 행으로 지정)에서는
>   해당 시트가 오류 없이 빈 결과(빈 컬렉션)로 처리된다(헤더 행·필수 컬럼이 유효하면 예외 발생 없음).
> - 단, `importLastDataRowIndex`를 `importFirstDataRowIndex`보다 작게 명시적으로 지정한 직접 역전은 `PxlDataException`을 발생시킨다.

---

## 셀 스타일러

셀 스타일은 컬럼 → 시트 → 워크북 → 내장 기본값 순으로 위임된다.
각 단계에서 지정하지 않았거나 적용할 수 없는 스타일러이면 다음 단계로 내려간다.

| 용도          | 기본 Styler                          |
|-------------|------------------------------------|
| 필수적인 헤더     | `PxlHeaderRequiredStyler`          |
| 선택적인 헤더     | `PxlHeaderOptionalStyler`          |
| 데이터         | `PxlDataVerticalCenterTextStyler`  |

- 내장 헤더 스타일러: `PxlHeaderRequiredStyler` · `PxlHeaderOptionalStyler` · `PxlHeaderHorizontalCenterTextStyler` ·
  `PxlHeaderVerticalCenterTextStyler` · `PxlHeaderWrapTextStyler`.
- 내장 데이터 스타일러: `PxlDataTextStyler` · `PxlDataThinBorderStyler` · `PxlDataVerticalCenterTextStyler` ·
  `PxlDataHorizontalCenterTextStyler` · `PxlDataWrapTextStyler` · `PxlDataCommaSeparatedNumericStyler`.
- 직접 구현하려면 `PxlStyler`(`Font apply(Workbook, CellStyle)`)를 구현한다.  
  `CellStyle`을 수정하고 셀에 적용할 `Font`를 반환한다. (예제는 [컬럼: 셀 스타일러 Export](#컬럼-셀-스타일러-export) 참고).

---

## 다양한 예제

### 컬럼: 날짜·숫자 형식

```java
@PxlColumn(name = "HireDate", pattern = "yyyy/MM/dd")
private LocalDate hireDate;

@PxlColumn(name = "StartTime", pattern = "HH:mm")
private LocalTime startTime;

@PxlColumn(name = "Amount", pattern = "#,##0.00")   // 수치는 DecimalFormat
private BigDecimal amount;
```

> export/import에서 형식을 다르게 하려면 `exportPattern` / `importPattern`을 따로 준다.  
> 날짜·시각은 export 시에는 형식이 강제되고, import 시에는 엄격 강제되지 않는다(고정 ISO 기본 패턴으로 파싱).  
> 기본 write 패턴은 ISO-8601이라 패턴 없는 값의 출력은 어느 머신에서나 동일하다.  
> LocalDateTime의 기본 read 패턴은 `T`(`yyyy-MM-dd'T'HH:mm:ss`)와 공백(`yyyy-MM-dd HH:mm:ss`)을 모두 받는다. 

### 컬럼: 참/거짓 문자열

```java
@PxlColumn(name = "Flag",
        exportTrueString = "Y", exportFalseString = "N",
        importTrueString = "Y", importFalseString = "N")   // 왕복하려면 export/import 같은 값
private Boolean flag;
```

### 컬럼: Collection

```java
@PxlColumn(name = "Tags", collectionSeparator = ",")   // 기본 구분자는 ";"
private List<String> tags;   // "a,b,c" ↔ [a, b, c]
```

> 빈/`null` 요소의 위치(인덱스)가 보존된다(`"a;;b"` ↔ `["a", null, "b"]`).  
> 요소 타입은 구체 클래스여야 하며 중첩 제네릭(`List<List<..>>`)·와일드카드는 지원하지 않는다.

### 컬럼: Enum

```java
@PxlColumn(name = "Grade")
private Grade grade;   // toString() 오버라이드값 또는 상수명과 매칭(대소문자·공백 무시)
```

### 컬럼: 커스텀 객체

`@PxlExportConverter`(인스턴스 또는 static, 반환 `String`) / `@PxlImportConverter`(static, 반환 타입=대상 타입)를 붙인다.

```java
@Getter                     // (선택) 사용자 편의용 — PXL은 아래 컨버터로만 변환하기 때문에 Getter가 필요하지는 않다.
@Setter                     // (선택) 사용자 편의용 — PXL은 아래 컨버터로만 변환하기 때문에 Setter가 필요하지는 않다.
@NoArgsConstructor          // (선택) PXL은 커스텀 객체를 무인자 생성자로 만들지 않는다.
@AllArgsConstructor         // 아래 parse()가 쓴다 — PXL 자체가 요구하는 것은 아니다.
public class Money {

    private String currency;
    private long amount;

    @PxlImportConverter                       // String → Money (반드시 static)
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

> 컨버터가 없으면 export는 재정의된 `toString()`, import는 `String` 단일인자 생성자로 변환한다.

### 컬럼: 마스킹 Export

```java
// 정규식에 매칭되는 문자를 '*'로 치환하여 export (마스킹된 값이 그대로 왕복)
@PxlColumn(name = "Secret", exportMasking = "\\d")     // 모든 숫자 마스킹
private String secret;
```

### 컬럼: 드롭다운 Export

```java
// 고정 목록 드롭다운
@PxlColumn(name = "Choice", exportOptionItems = {"Red", "Green", "Blue"})
private String choice;
```

### 컬럼: null Export

```java
// null을 특정 문자열로 export (기본값은 빈 문자열 "")
@PxlColumn(name = "Memo", exportNullString = "-")
private String memo;
```

### 컬럼: 수식 Export

```java
// 선두 '='가 붙은 문자열을 수식으로 계산해 셀에 적용 (비스트리밍 import 시 계산 결과가 읽힘)
@PxlColumn(name = "Total", exportStringAsFormula = true)
private String total;   // 예: "=A2+B2"
```

### 컬럼: 이미지 export

```java
// 문자열(이미지 URL/경로)을 실제 이미지로 삽입
@PxlColumn(name = "Photo", exportStringAsPicture = true)
private String photo;

// 한 셀에 여러 이미지 (List<String>)
@PxlColumn(name = "Gallery", exportStringAsPicture = true)
private List<String> gallery;
```

### 컬럼: 셀 스타일러 Export

스타일러는 컬럼 → 시트 → 워크북 → 내장 기본값 순으로 위임된다.  
내장 스타일러를 쓰거나 직접 구현한다(계단식·내장 목록은 [셀 스타일러](#셀-스타일러) 참고).

```java
// 컬럼에 내장 스타일러 지정 (데이터 셀 + 필수 헤더 셀)
@PxlColumn(name = "Amount",
        exportColumnDataCellStyler = PxlDataCommaSeparatedNumericStyler.class,
        exportColumnRequiredHeaderCellStyler = PxlHeaderHorizontalCenterTextStyler.class)
private BigDecimal amount;
```

```java
// 워크북 전체 데이터 셀 스타일 (옵션) — 시트/컬럼에서 따로 지정 안 하면 이게 적용
PxlExportWorkbookOption.builder()
                       .exportWorkbookDataCellStyler(PxlDataThinBorderStyler.class)
                       .build();
```

```java
// 직접 만들기 — PxlStyler 구현: CellStyle을 수정하고, 셀에 적용할 Font를 반환한다
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
        return font;                       // 반환한 Font가 셀에 적용된다
    }
}
// @PxlColumn(name = "X", exportColumnDataCellStyler = BoldCenterStyler.class)
```

### 컬럼: 열 순서 고정 Export

필드 선언 순서가 열 순서를 보장하지는 않는다. 순서가 중요하면 `exportOrder`를 설정한다.

```java
@PxlColumn(name = "A", exportOrder = "01")
private String a;

@PxlColumn(name = "B", exportOrder = "02")
private String b;
```

> `exportOrder`는 문자열 사전식 비교다.
> 숫자 순서를 원하면 `"01"`, `"02"`, … 처럼 0으로 자릿수를 채운다(`"2"` vs `"10"` → `"10"`이 먼저 옴).

### 시트: 다중 시트 Export

```java
// 호출 순서 = 시트 순서(Engineering, Sales, Departments), 시트마다 행 클래스를 달리 줄 수 있다
pxl.exportExcel()
   .sheet(Employee.class, engineering, "Engineering")
   .sheet(Employee.class, sales, "Sales")
   .sheet(Department.class, departments, "Departments")
   .toFile(new File("company.xlsx"));

// 샘플 export도 동일하다
pxl.exportSampleExcel()
   .sheet(Employee.class, "Employees")
   .sheet(Department.class, "Departments")
   .toFile(new File("template.xlsx"));
```

### 시트: 다중 시트 Import

```java
import io.github.hclimkr.pxl.builder.PxlExcelImportBuilder;

// 1) 워크북 형태: @PxlSheet 필드별로 여러 시트를 한 번에 바인딩
//    (Company = @PxlWorkbook 클래스, employees/departments 두 개의 @PxlSheet 필드를 가진다)
Company company = pxl.importExcel()
                     .workbook(Company.class)
                     .fromFile(file);

// 2) 시트 형태: 같은 빌더로 시트마다 호출해 각각 실행한다(시트별로 반환 타입이 다를 수 있다)
PxlExcelImportBuilder builder = pxl.importExcel();
List<Employee> employees = builder.sheet(Employee.class, "Employees")
                                  .fromFile(file);
List<Department> departments = builder.sheet(Department.class, "Departments")
                                      .fromFile(file);
```

### 시트: 컬럼 값으로 그룹핑하여 시트 분할 Export

`@PxlSheet(exportGroupingFieldName = "필드명")`을 주면 그 값별로 여러 시트로 나뉜다.

```java
@PxlSheet(name = "Employees", exportGroupingFieldName = "department")
private List<Employee> employees;   // department 값마다 별도 시트로 export
```

### 시트: 데이터가 1행부터 시작하지 않을 때 Import

헤더·데이터의 시작 행/열을 1-based로 지정한다.
워크북 형태는 `@PxlSheet`에 직접 주고, 시트 형태는 `PxlImportSheetOption`으로 준다.

```java
// 워크북 클래스: @PxlSheet에 지정 (헤더 3행, 데이터 4행부터)
@PxlSheet(name = "Employees", importHeaderRowIndex = 3, importFirstDataRowIndex = 4)
private List<Employee> employees;
```

```java
// 시트 형태: 행/열 위치는 PxlImportSheetOption에 있고, 그것을 워크북 옵션의 importSheetOptions에 담는다
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

> 끝 위치·열 범위(`importLastDataRowIndex`, `importFirstDataColumnIndex`, `importLastDataColumnIndex`)도 같은 방식으로 준다. 모두 1-based.
> Streaming Reader는 `getFirstRowNum()`이 없으므로 헤더 행 지정이 필수다.

### 시트: CSV 인코딩·구분자 Import

```java
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

List<Employee> rows = pxl.importCsv()
                         .override(PxlImportWorkbookOption.builder()
                                                          .importCsvCharset("US-ASCII") // 기본값은 "UTF-8" (레거시 인코딩만 지정)
                                                          .importCsvDelimiter('\t')     // char (예: TSV). 기본 ','
                                                          .build())
                         .sheet(Employee.class)
                         .fromFile(file);
```

> CSV 기본 인코딩은 `UTF-8`이다.
> 다른 인코딩(예: `US-ASCII`·`MS949`·`EUC-KR`) CSV는 위처럼 `importCsvCharset(...)`으로 지정한다. `importCsvDelimiter`는 `char`이므로 작은따옴표를 쓴다.

CSV 워크북은 시트 하나가 파일 하나로 읽히므로 시트끼리 인코딩·구분자가 같을 필요가 없다.  
그래서 두 속성은 `@PxlSheet`에도 있으며, 둘 다 지정하지 않은 시트는 워크북 값을 상속한다.

```java
@PxlWorkbook(importCsvCharset = "MS949")            // 따로 지정하지 않은 모든 시트에 적용
public class CompanyWorkbook {

    @PxlSheet(name = "Legacy")                       // MS949 상속
    private List<CharsetRow> legacy;

    @PxlSheet(name = "Modern", importCsvCharset = "UTF-8")   // 이 파일만 UTF-8로 읽음
    private List<CharsetRow> modern;
}

CompanyWorkbook workbook = pxl.importCsv()
                              .workbook(CompanyWorkbook.class)
                              .fromFiles(Arrays.asList(legacyCsv, modernCsv));
```

해석은 위에서 아래로 내려가며, 값을 지정한 첫 단계가 우선순위 높다.

| 단계 | "미지정"을 뜻하는 값 |
|------|--------------------|
| `PxlImportSheetOption.importCsvCharset` / `importCsvDelimiter` | `null`, 그리고 `""` / `'\0'`도 동일 |
| `@PxlSheet(importCsvCharset)` / `(importCsvDelimiter)`         | `""` / `'\0'`      |
| `PxlImportWorkbookOption.importCsvCharset` / `importCsvDelimiter` | `null`, 그리고 `""` / `'\0'`도 동일 |
| `@PxlWorkbook(importCsvCharset)` / `(importCsvDelimiter)`      | `""` / `'\0'`      |
| 내장 기본값                                                     | `"UTF-8"` / `','`  |

> "미지정"이 실효 기본값이 아니라 미지정 표식이므로, 워크북이 다른 값을 지정한 상황에서 시트가 `"UTF-8"`·`','`를 명시해 기본값으로 되돌릴 수 있다.  
> 시트 폼(`sheet(...)`)에는 `@PxlSheet`를 붙일 필드가 없으므로 와일드카드 `PxlImportSheetOption`이 유일한 시트 단계 경로다 — *옵션 구조* 참고.  
> 한 단계 위도 마찬가지다. 시트 폼은 워크북 클래스 없이 워크북 메타를 만들므로 `@PxlWorkbook`도 읽지 않는다. 시트 폼에서는 애노테이션 두 단계가 모두 무효이고 옵션만 적용된다.

### 시트: CSV 인코딩·구분자·BOM Export

쓰는 쪽은 읽는 쪽과 대칭이며 BOM이 하나 더 붙는다 — CSV는 시트 하나가 파일 하나로 기록되므로 셋 다 스키마가 아니라 파일의 속성이다.

```java
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;

pxl.exportCsv()
   .sheet(Employee.class, employees, "Employees")
   .override(PxlExportWorkbookOption.builder()
                                    .exportCsvCharset("MS949")   // 기본값은 "UTF-8"
                                    .exportCsvDelimiter('\t')    // char (예: TSV). 기본값 ','
                                    .exportCsvBom(true)          // 옵션은 Boolean, 애노테이션은 PxlOptionalBoolean
                                    .build())
   .toFile(csvFile);
```

```java
@PxlWorkbook(exportCsvCharset = "MS949", exportCsvBom = PxlOptionalBoolean.TRUE)
public class CompanyWorkbook {

    @PxlSheet(name = "Legacy")                                    // MS949와 BOM을 상속
    private List<CharsetRow> legacy;

    @PxlSheet(name = "Modern",
              exportCsvCharset = "UTF-8",
              exportCsvBom = PxlOptionalBoolean.FALSE)            // 이 파일만 BOM 없이 쓴다
    private List<CharsetRow> modern;
}
```

해석 순서는 import 쪽과 같은 5단이다.

| 단계 | "미지정"을 뜻하는 값 |
|------|--------------------|
| `PxlExportSheetOption.exportCsvCharset` / `exportCsvDelimiter` / `exportCsvBom` | `null`, 그리고 `""` / `'\0'`도 동일 |
| `@PxlSheet(exportCsvCharset)` / `(exportCsvDelimiter)` / `(exportCsvBom)`       | `""` / `'\0'` / `UNSPECIFIED` |
| `PxlExportWorkbookOption.exportCsvCharset` / `exportCsvDelimiter` / `exportCsvBom` | `null`, 그리고 `""` / `'\0'`도 동일 |
| `@PxlWorkbook(exportCsvCharset)` / `(exportCsvDelimiter)` / `(exportCsvBom)`    | `""` / `'\0'` / `UNSPECIFIED` |
| 내장 기본값                                                                      | `"UTF-8"` / `','` / `false` |

> 애노테이션에서 평범한 `boolean`을 못 쓰는 것은 `exportCsvBom`뿐이다. `boolean`에는 "미지정"에 쓸 값이 남지 않아 `false`가 침묵과 구분되지 않고,
> 그러면 워크북이 켠 BOM을 시트가 끌 방법이 없다. 그 세 번째 값을 채워 주는 것이 `PxlOptionalBoolean`(`UNSPECIFIED` · `TRUE` · `FALSE`)이며,
> 옵션 클래스는 `null`이 이미 그 뜻이라 박싱된 `Boolean`을 그대로 쓴다.  
> BOM 자체는 UTF-8·UTF-16LE·UTF-16BE에서만 기록된다 — [제약 (Limitation)](#제약-limitation) 참고.  
> import 쪽과 마찬가지로 시트 폼(`sheet(...)`)은 애노테이션 두 단계를 모두 읽지 않으므로 그곳에서는 옵션이 경로다.

### 워크북: 비밀번호로 암호화 Export

```java
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;

pxl.exportExcel()
   .sheet(Employee.class, employees, "Employees")
   .override(PxlExportWorkbookOption.builder()
                                    .exportPassword("secret")
                                    .build())
   .toFile(new File("secured.xlsx"));
```

### 워크북: Streaming Reader

```java
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

// 스트리밍은 헤더 행 위치를 정확히 지정해야 하므로 시트 옵션과 함께 준다
PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                                                       .importHeaderRowIndex(1)   // 1-based
                                                       .build();

List<Employee> rows = pxl.importExcel()
                         .override(PxlImportWorkbookOption.builder()
                                                          .importUsingStreamReader(true)   // XLSX 전용
                                                          .importSheetOptions(Arrays.asList(sheetOption))
                                                          .build())
                         .sheet(Employee.class, "Employees")
                         .fromFile(bigFile);
```

> Streaming Reader는 XLSX 전용이고 수식 셀을 평가하지 못하며, `getFirstRowNum()`이 없어 헤더 행을 정확히 지정해야 한다.

---

## 제약 (Limitation)

| 형식   | 최대 시트   | 최대 행      | 최대 열   | 비고                                             |
|------|---------|-----------|--------|------------------------------------------------|
| XLSX | 100     | 1,048,576 | 16,384 | Streaming Reader로 메모리 문제(GC overhead) 없이 읽기 가능 |
| XLS  | 100     | 65,536    | 256    | Streaming Reader 미지원이나, 비스트리밍으로도 메모리 문제 없음     |
| CSV  | 100     | 100,000   | 16,384 | 파일 하나가 시트 하나이므로 "시트"는 `fromFiles(...)`/`fromStreams(...)`에 넘긴 파일 수를 뜻한다 — CSV export는 시트 하나를 쓰므로 그쪽에는 행/열 한도만 적용된다.<br/>열 한도는 XLSX와 같게 맞춰 XLSX로 export되는 행 클래스가 CSV로도 그대로 읽히게 했고, 행 한도가 더 낮은 것은 import 시 CSV를 통째로 메모리에 올린 뒤에야 한도를 검사하기 때문이다 |

> 이 한도는 `PxlFileFormat`이 갖고 있는 값(`getMaxExportRows()` 등)이다. 따라서 엔진은 자기가 쓰는 형식의 한도를 그대로 따르며, `XSSF`와 `SXSSF`도 마찬가지다.  
> XLSX·XLS의 행/열 수치는 형식 자체의 한도지만, 시트 수와 CSV의 모든 수치는 형식이 아니라 PXL이 두는 한도다. 어느 쪽이든 초과하면 `PxlDataException`이다.

- 자동 열 너비(`autoSizeColumn`)와 대량 데이터  
  `exportColumnWidth`를 지정하지 않으면 기본값(auto)이 적용되어, export 시 POI `autoSizeColumn`이 해당 열의 모든 행 셀을 폰트 메트릭으로 실측한다(열당 O(행 수)).  
  컬럼 N·행 M이면 O(N×M)의 측정 비용이 추가되어 대량 데이터 export에서 성능 저하의 지배적 요인이 될 수 있다.  
  행 수가 많으면 `@PxlColumn(exportColumnWidth = ...)` 또는 옵션으로 고정 너비를 지정하는 것을 권장한다.

- export Heap 사용량과 `SXSSF`  
  `XSSF`(기본값)는 한 바이트를 쓰기 전에 워크북 전체를 객체 그래프로 구성하므로, 최종 생성되는 파일보다 훨씬 큰 Heap을 점유한다.  
  `SXSSF`는 같은 `.xlsx`를 만들면서 행을 슬라이딩 윈도우(`exportSXSSFRowAccessWindowSize`, 기본 100)만큼만 메모리에 두고 나머지는 임시 파일로 내보내므로, 힙 사용량이 행 수와 거의 무관해진다.  
  XLSX 전용이라 `HSSF` 엔진(`.xls`)에는 효과가 없다.  
  단, 자동 너비로 둔 열은 실측을 위해 추적 대상이 되고 추적된 열은 메모리에 남으므로 절감 효과를 깎아먹는다. `SXSSF`를 쓸 때는 `exportColumnWidth`를 고정하는 편이 좋다.

- CSV export는 쓰기 전에 출력물을 만들며, 크면 임시 파일로 넘긴다  
  CSV export는 목적지를 열기 전에 파일 전체를 렌더링한다. 그래서 실패해도 파일이 남지 않고 `toStream(...)`이 한 번에 쓸 수 있다.
  `PxlConstants.EXPORT_MEMORY_THRESHOLD_OF_CSV`(4 MiB)까지는 그 출력물을 메모리에 두고, 그 이상은 임시 파일로 이어 쓰므로 **필요한 힙이 출력 크기를 따라 늘지 않는다.**
  임시 파일은 `java.io.tmpdir` 아래 `pxl-csv-export-` 접두어로 만들어지고, 성공·실패와 무관하게 호출이 끝나기 전에 지워진다.
  이 spill에서 알아 둘 것이 셋이다. 큰 export는 **디스크 여유 공간**이 필요하며, 모자라면 `PxlIOException`으로 실패한다. 임시 파일은 **평문으로 기록**되는데, CSV export가 `exportPassword`를 암호화 대신 거부하기 때문에 더 유의해야 한다 — 민감한 데이터라면 `java.io.tmpdir`를 신뢰할 수 있는 위치로 지정한다. 그리고 export 도중 **JVM이 강제 종료되면 파일이 남는다**(정상 반환이나 예외만이 정리 단계에 도달한다).
  그럼에도 메모리가 모자라면 `OutOfMemoryError`가 `PxlException`이 아니라 그 자체로 올라온다 — `Error`는 `Exception`이 아니라서 마지막 단계의 정규화가 덮지 못하며, 잡는 것도 옳지 않다(감싸는 행위 자체가 메모리가 없어서 난 상황에서 다시 할당하는 일이다). 넘긴 행 객체는 여전히 전량 상주하므로, 힙은 그 쪽을 기준으로 잡는다.

- CSV export가 무시하는 설정  
  셀의 겉모습이나 워크북의 구성만 정하는 것들이다 — `@PxlColumn(exportColumn*Styler, exportColumnWidth, exportOptionItems, exportEnumDropDownListStyle)`,
  `@PxlSheet(exportSheet*Styler, exportRowHeightInPoints, exportColumnFilter, exportGroupingFieldName)`,
  `@PxlWorkbook(exportWorkbook*Styler, exportExcelEngine, exportSXSSFRowAccessWindowSize)`, 그리고 헤더 창 고정.
  그룹핑이 무시된다는 것은 행이 그룹별로 모이지 않고 넘긴 순서 그대로 기록된다는 뜻이다.  
  다만 둘은 값을 버리지 않고 파일에 남긴다 — `exportStringAsFormula`는 앞의 `=`까지 포함해 텍스트를 그대로 쓰고(평가하지 않는다),
  `exportStringAsPicture`는 그림을 박는 대신 이미지 경로를 쓴다. 사라질 줄 알았던 경로·URL이 노출되는 셈이다.  

- CSV BOM  
  `exportCsvBom`은 UTF-8·UTF-16LE·UTF-16BE에서만 적용된다. 그 외 문자셋에서는 켜 두어도 예외도 경고도 없이 생략된다 —
  `UTF-16`은 인코더가 이미 하나를 붙이고, EUC-KR 같은 비유니코드 문자셋은 U+FEFF를 인코딩하지 못해 `?`로 치환되어 첫 헤더 필드를 깨뜨리기 때문이다.
  런타임에 알아챌 방법이 없으므로, BOM이 빠졌다면 먼저 문자셋을 확인한다.

---

## 자주 겪는 함정 체크리스트

- ✅ **DTO에 무인자 생성자가 필요하다.**  
  import 시 리플렉션으로 행 객체를 만든다.  
  Lombok을 쓸 때 `@AllArgsConstructor`만 붙이면 무인자 생성자가 사라지므로 `@NoArgsConstructor`를 함께 붙인다.
- ✅ **이름 매칭**  
  `@PxlColumn`/`@PxlSheet`의 `name`(미지정 시 필드명)이 실제 헤더/시트명과 일치해야 한다. 둘 다 공백은 무시한다. 시트명은 대소문자도 무시하지만 컬럼 헤더는 대소문자를 구분한다.
- ✅ **이름 안 맞는 컬럼**  
  필수(`@NotNull`/`@NotEmpty`/`@NotBlank`)면 예외 발생, 필수가 아니면 조용히 제외된다.
- ✅ **인덱스는 1-based**  
  `exportFirstDataColumnIndex`, `importHeaderRowIndex` 등 모두 1-based이다.  
  `@PxlRowIndex`가 받는 값도 마찬가지로 1-based로, 가져온 행의 스프레드시트 행 번호이다.
- ✅ **CSV 기본 인코딩은 `UTF-8`**  
  다른 인코딩(`US-ASCII`·`MS949`·`EUC-KR` 등)은 쓸 때 `exportCsvCharset(...)`으로, 읽을 때 `importCsvCharset(...)`으로 명시한다.
  워크북 폼에서는 시트마다 파일이 다르므로, 하나만 다른 시트는 워크북 전체를 한 설정으로 몰지 말고 `@PxlSheet(exportCsvCharset)` / `(exportCsvDelimiter)` / `(exportCsvBom)` — 읽는 쪽은 `(importCsvCharset)` / `(importCsvDelimiter)` — 로 그 시트만 지정한다.
- ✅ **CSV export는 필요한 값만 인용한다**  
  `"010"`·`"1E+10"` 같은 값은 인용 없이 기록되므로 Excel로 열면 숫자 `10`·`1e10`으로 재해석된다. PXL로 다시 읽으면 원래 문자열 그대로다 — 손실은 파일이 아니라 스프레드시트 앱에서 생긴다.
  아직 인용 정책 설정은 없으므로, 표시가 중요하면 `pattern`으로 써 둔다.
- ✅ **`long` / `BigInteger` / `BigDecimal` 정밀도**  
  큰 수는 숫자 셀(double, 2^53 한계)에서 정밀도가 손실될 수 있다. 정확히 보존하려면 `pattern`으로 문자열 셀 출력하거나 `BigInteger`/`BigDecimal`을 쓴다.
- ✅ **`Pxl`은 재사용**한다.  
  `new Pxl()`은 유효성 검사 부트스트랩 비용이 있으니 싱글톤/스프링 빈으로 둔다(thread-safe).
- ✅ **Streaming Reader 제약**  
  XLSX 전용, 수식 평가 불가, 헤더 행 위치를 정확히 지정해야 한다.
- ✅ **대량 export 성능**  
  열 너비 미지정 시 `autoSizeColumn`이 전체 행을 실측한다. 행이 많으면 `@PxlColumn(exportColumnWidth = ...)`로 고정 너비를 준다.

---

## 라이선스

이 프로젝트는 **[Apache License 2.0](../LICENSE)** 하에 배포된다(전문은 [README의 라이선스](../README_ko.md#라이선스)를 참고한다).
