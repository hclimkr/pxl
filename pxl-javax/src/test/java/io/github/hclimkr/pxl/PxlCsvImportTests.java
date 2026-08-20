package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.builder.PxlCsvImportBuilder;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.IllegalCharsetNameException;
import java.nio.charset.StandardCharsets;
import java.nio.charset.UnsupportedCharsetException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.time.LocalDate;
import java.util.*;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * CSV import tests. (Since CSV supports import only, this verifies import rather than round-trip.)
 * <p>
 * CSV sources are generated in English (ASCII) within the tests. ASCII bytes are interpreted identically under the default encoding (UTF-8).
 * Covers the 4 CSV import methods of {@link Pxl} (file single/multi, stream single/multi).
 */
public class PxlCsvImportTests {

    private static Pxl pxl;

    // Headers must match @PxlColumn(name=...). Active is written with the default true/false strings (yes/no).
    private static final String EMPLOYEES_CSV =
            "Name,Age,Salary,Active,HireDate,Grade,Department\n" +
                    "Alice,30,50000.50,yes,2020-01-15,A,Engineering\n" +
                    "Bob,42,72000.00,no,2018-07-01,B,Sales\n";

    private static final String DEPARTMENTS_CSV =
            "Code,DepartmentName,Headcount\n" +
                    "ENG,Engineering,12\n" +
                    "SAL,Sales,7\n";

    // Tab-delimited CSV (for verifying the importCsvDelimiter option)
    private static final String EMPLOYEES_TSV =
            "Name\tAge\tSalary\tActive\tHireDate\tGrade\tDepartment\n" +
                    "Alice\t30\t50000.50\tyes\t2020-01-15\tA\tEngineering\n";

    private static final String DEPARTMENTS_TSV =
            "Code\tDepartmentName\tHeadcount\n" +
                    "ENG\tEngineering\t12\n";

    // Non-ASCII payloads for the charset cascade. Only the encoding is under test, so the text itself is incidental -
    // what matters is that each string is unreadable under the wrong charset of the pair.
    private static final String SEOUL_CSV = "City\n서울\n";
    private static final String BUSAN_CSV = "City\n부산\n";
    private static final String CAFE_CSV = "City\ncafé\n";

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    private static InputStream stream(final String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.US_ASCII));
    }

    private static InputStream stream(final String content, final String charsetName) {
        return new ByteArrayInputStream(content.getBytes(Charset.forName(charsetName)));
    }

    private static File writeCsv(final Path dir, final String fileName, final String content) throws Exception {
        final Path path = dir.resolve(fileName);
        Files.write(path, content.getBytes(StandardCharsets.US_ASCII));
        return path.toFile();
    }

    private static void assertEmployees(final List<Employee> employees) {
        assertThat(employees).hasSize(2);

        final Employee alice = employees.get(0);
        assertThat(alice.getName()).isEqualTo("Alice");
        assertThat(alice.getAge()).isEqualTo(30);
        assertThat(alice.getSalary()).isEqualByComparingTo("50000.50");
        assertThat(alice.getActive()).isTrue();
        assertThat(alice.getHireDate()).isEqualTo(LocalDate.of(2020, 1, 15));
        assertThat(alice.getGrade()).isEqualTo(Grade.A);
        assertThat(alice.getDepartment()).isEqualTo("Engineering");

        final Employee bob = employees.get(1);
        assertThat(bob.getName()).isEqualTo("Bob");
        assertThat(bob.getActive()).isFalse();
        assertThat(bob.getGrade()).isEqualTo(Grade.B);
    }

    // ------------------------------------------------------------------
    // Sheet form (single CSV -> Collection<Row>)
    // ------------------------------------------------------------------

    @Test
    public void importCsvFile_singleSheet_binds(@TempDir final Path tempDir) throws Exception {
        final File csvFile = writeCsv(tempDir, "Employees.csv", EMPLOYEES_CSV);

        final List<Employee> employees = pxl.importCsv()
                .sheet(Employee.class)
                .fromFile(csvFile);

        assertEmployees(employees);
    }

    @Test
    public void importCsvStream_singleSheet_binds() throws Exception {
        final List<Employee> employees = pxl.importCsv()
                .sheet(Employee.class)
                .fromStream("Employees", stream(EMPLOYEES_CSV));

        assertEmployees(employees);
    }

    // Every CSV field arrives as a string, so this drives the string dispatcher rather than the cell one.
    @Test
    public void importCsvStream_uuidColumns_bind() throws Exception {
        final String uuidText = "123e4567-e89b-12d3-a456-426614174000";
        final String otherUuidText = "00112233-4455-6677-8899-aabbccddeeff";
        final String csv = "Id,Ids\n"
                + uuidText + "," + otherUuidText + ";" + uuidText + "\n";

        final UuidRow row = pxl.importCsv()
                .sheet(UuidRow.class)
                .fromStream("Uuids", stream(csv)).get(0);

        assertThat(row.getId()).isEqualTo(UUID.fromString(uuidText));
        assertThat(row.getIds()).containsExactly(UUID.fromString(otherUuidText), UUID.fromString(uuidText));
    }

    @Test
    public void importCsvStream_uuidShortGroups_throws() throws Exception {
        // The canonical form is enforced on this path too, not only when reading a cell.
        final String csv = "Id\n1-1-1-1-1\n";

        assertThrows(PxlCellCodecException.class, () -> pxl.importCsv()
                .sheet(UuidRow.class)
                .fromStream("Uuids", stream(csv)));
    }

    // Resource ownership: fromStream does not close the InputStream passed by the caller
    @Test
    public void importCsvStream_doesNotCloseCallerStream() throws Exception {
        final boolean[] closed = {false};
        final ByteArrayInputStream tracking = new ByteArrayInputStream(EMPLOYEES_CSV.getBytes(StandardCharsets.US_ASCII)) {
            @Override
            public void close() {
                closed[0] = true;
            }
        };

        final List<Employee> employees = pxl.importCsv()
                .sheet(Employee.class)
                .fromStream("Employees", tracking);

        assertThat(employees).hasSize(2);
        assertThat(closed[0]).as("fromStream must not close the caller's stream").isFalse();
    }

    @Test
    public void importCsvStream_singleSheet_intoSet_returnsSet() throws Exception {
        // collectionClass=Set.class -> the return must be a Set implementation (the default sheet() returns a List).
        @SuppressWarnings("unchecked") final Set<Employee> employees =
                pxl.importCsv()
                        .sheet(Employee.class, Set.class)
                        .fromStream("Employees", stream(EMPLOYEES_CSV));

        assertThat(employees).isInstanceOf(Set.class);
        assertThat(employees).isNotInstanceOf(List.class);
        assertThat(employees).extracting(Employee::getName).containsExactlyInAnyOrder("Alice", "Bob");
    }

    @Test
    public void importCsvStream_tabDelimiter_binds() throws Exception {
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvDelimiter('\t')
                .build();

        final List<Employee> employees = pxl.importCsv()
                .override(option)
                .sheet(Employee.class)
                .fromStream("Employees", stream(EMPLOYEES_TSV));

        assertThat(employees).hasSize(1);
        assertThat(employees.get(0).getName()).isEqualTo("Alice");
        assertThat(employees.get(0).getSalary()).isEqualByComparingTo("50000.50");
    }

    @Test
    public void importCsv_unsupportedCharset_throwsArgumentException() {
        // Charset.forName throws unchecked, so this used to escape the IOException-only try and reach the builder
        // boundary as a PxlSystemException that named neither the attribute nor its value.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("NoSuchCharset-1")
                .build();

        final PxlArgumentException exception = assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .override(option)
                .sheet(Employee.class)
                .fromStream("Employees", stream(EMPLOYEES_CSV)));

        // Both bundles name the attribute and echo the value, so these hold whatever the process locale is.
        assertThat(exception).hasMessageContaining("importCsvCharset");
        assertThat(exception).hasMessageContaining("NoSuchCharset-1");
        assertThat(exception).hasCauseInstanceOf(UnsupportedCharsetException.class);
    }

    @Test
    public void importCsv_malformedCharsetName_throwsArgumentException() {
        // A name that is not merely unsupported but illegally formed takes the other Charset.forName branch.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("UTF 8")
                .build();

        final PxlArgumentException exception = assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .override(option)
                .sheet(Employee.class)
                .fromStream("Employees", stream(EMPLOYEES_CSV)));

        assertThat(exception).hasMessageContaining("importCsvCharset");
        assertThat(exception).hasCauseInstanceOf(IllegalCharsetNameException.class);
    }

    @Test
    public void importCsv_lineBreakDelimiter_throwsArgumentException() {
        // CSVFormat rejects the delimiter while the format is built, also unchecked.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvDelimiter('\n')
                .build();

        final PxlArgumentException exception = assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .override(option)
                .sheet(Employee.class)
                .fromStream("Employees", stream(EMPLOYEES_CSV)));

        assertThat(exception).hasMessageContaining("importCsvDelimiter");
        assertThat(exception).hasCauseInstanceOf(IllegalArgumentException.class);
    }

    @Test
    public void importCsv_quoteCharAsDelimiter_throwsArgumentException() {
        // The other rejected delimiter: identical to the quote character.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvDelimiter('"')
                .build();

        final PxlArgumentException exception = assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .override(option)
                .sheet(Employee.class)
                .fromStream("Employees", stream(EMPLOYEES_CSV)));

        assertThat(exception).hasMessageContaining("importCsvDelimiter");
        assertThat(exception).hasCauseInstanceOf(IllegalArgumentException.class);
    }

    @Test
    public void importCsvStream_overrideAfterSheetConfig_binds() throws Exception {
        // Same option as importCsvStream_tabDelimiter_binds, but chained after sheet(...) on the source step.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvDelimiter('\t')
                .build();

        final List<Employee> employees = pxl.importCsv()
                .sheet(Employee.class)
                .override(option)
                .fromStream("Employees", stream(EMPLOYEES_TSV));

        assertThat(employees).hasSize(1);
        assertThat(employees.get(0).getName()).isEqualTo("Alice");
        assertThat(employees.get(0).getSalary()).isEqualByComparingTo("50000.50");
    }

    @Test
    public void importCsv_sheetCalledTwicePerBuilder_bindsEachSourceIndependently() throws Exception {
        // Unlike the export builder, sheet(...) returns a source step rather than the builder itself, so the calls
        // are not chained; the same builder instance is reused once per CSV source, each with its own row class.
        final PxlCsvImportBuilder builder = pxl.importCsv();

        final List<Employee> employees = builder
                .sheet(Employee.class)
                .fromStream("Employees", stream(EMPLOYEES_CSV));
        final List<Department> departments = builder
                .sheet(Department.class)
                .fromStream("Departments", stream(DEPARTMENTS_CSV));

        // The second call must not be affected by the first (no state leaks between source steps).
        assertEmployees(employees);
        assertThat(departments).extracting(Department::getCode).containsExactly("ENG", "SAL");
        assertThat(departments.get(0).getHeadcount()).isEqualTo(12);
    }

    // ------------------------------------------------------------------
    // Workbook form (multiple CSVs -> @PxlWorkbook, with file name/name becoming the sheet name)
    // ------------------------------------------------------------------

    @Test
    public void importCsvFiles_multiSheetWorkbook_binds(@TempDir final Path tempDir) throws Exception {
        // The file name (without extension) matches the @PxlSheet name.
        final File employeesCsv = writeCsv(tempDir, "Employees.csv", EMPLOYEES_CSV);
        final File departmentsCsv = writeCsv(tempDir, "Departments.csv", DEPARTMENTS_CSV);

        final CompanyWorkbook workbook = pxl.importCsv()
                .workbookName("Acme")
                .workbook(CompanyWorkbook.class)
                .fromFiles(Arrays.asList(employeesCsv, departmentsCsv));

        assertThat(PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(workbook)).isEqualTo("Acme");
        assertEmployees(workbook.getEmployees());

        final List<Department> departments = workbook.getDepartments();
        assertThat(departments).hasSize(2);
        assertThat(departments.get(0).getCode()).isEqualTo("ENG");
        assertThat(departments.get(0).getHeadcount()).isEqualTo(12);
    }

    @Test
    public void importCsvStreams_multiSheetWorkbook_binds() throws Exception {
        // csvNames match the @PxlSheet name.
        final List<String> csvNames = Arrays.asList("Employees", "Departments");
        final List<InputStream> csvStreams = Arrays.asList(stream(EMPLOYEES_CSV), stream(DEPARTMENTS_CSV));

        final CompanyWorkbook workbook = pxl.importCsv()
                .workbookName("Acme")
                .workbook(CompanyWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(workbook)).isEqualTo("Acme");
        assertEmployees(workbook.getEmployees());
        assertThat(workbook.getDepartments()).hasSize(2);
        assertThat(workbook.getDepartments().get(1).getDepartmentName()).isEqualTo("Sales");
    }

    @Test
    public void importCsvStreams_workbookNameAfterWorkbookConfig_setsName() throws Exception {
        // workbookName(...) chained after workbook(...) on the source step must set the @PxlWorkbookName field.
        final List<String> csvNames = Arrays.asList("Employees", "Departments");
        final List<InputStream> csvStreams = Arrays.asList(stream(EMPLOYEES_CSV), stream(DEPARTMENTS_CSV));

        final CompanyWorkbook workbook = pxl.importCsv()
                .workbook(CompanyWorkbook.class)
                .workbookName("Acme")
                .fromStreams(csvNames, csvStreams);

        assertThat(PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(workbook)).isEqualTo("Acme");
        assertEmployees(workbook.getEmployees());
    }

    @Test
    public void importCsvFile_workbookForm_fileNameSelectsSheet(@TempDir final Path tempDir) throws Exception {
        // fromFile derives the sheet name from the file name (extension removed) and routes the CSV to the
        // @PxlSheet whose name equals that base name. The SAME workbook class binds a different sheet
        // depending only on the file name.
        final File employeesCsv = writeCsv(tempDir, "Employees.csv", EMPLOYEES_CSV);
        final File departmentsCsv = writeCsv(tempDir, "Departments.csv", DEPARTMENTS_CSV);

        // "Employees.csv" -> only the Employees sheet is populated
        final CompanyWorkbook byEmployees = pxl.importCsv()
                .workbook(CompanyWorkbook.class)
                .fromFile(employeesCsv);
        assertEmployees(byEmployees.getEmployees());
        assertThat(byEmployees.getDepartments()).isNullOrEmpty();

        // "Departments.csv" -> only the Departments sheet is populated (same class, different file name)
        final CompanyWorkbook byDepartments = pxl.importCsv()
                .workbook(CompanyWorkbook.class)
                .fromFile(departmentsCsv);
        assertThat(byDepartments.getEmployees()).isNullOrEmpty();
        assertThat(byDepartments.getDepartments()).hasSize(2);
        assertThat(byDepartments.getDepartments().get(0).getCode()).isEqualTo("ENG");
    }

    // ------------------------------------------------------------------
    // Encoding (importCsvCharset)
    // ------------------------------------------------------------------

    @Test
    public void importCsvStream_latin1Charset_decoded() throws Exception {
        // The é in "café" is a single byte 0xE9 in ISO-8859-1
        final byte[] csvBytes = "City\ncafé\n".getBytes(Charset.forName("ISO-8859-1"));

        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("ISO-8859-1")
                .build();

        final InputStream inputStream = new ByteArrayInputStream(csvBytes);
        final List<CharsetRow> rows = pxl.importCsv()
                .override(option)
                .sheet(CharsetRow.class)
                .fromStream("Cities", inputStream);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getCity()).isEqualTo("café");
    }

    @Test
    public void importCsvStream_ms949Charset_decodesKorean() throws Exception {
        // "서울" is 4 bytes in MS949 (2 bytes per Hangul char). Overriding the UTF-8 default with a legacy
        // Korean charset must decode it correctly; read under the UTF-8 default it would be garbled.
        final byte[] csvBytes = "City\n서울\n".getBytes(Charset.forName("MS949"));

        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("MS949")
                .build();

        final InputStream inputStream = new ByteArrayInputStream(csvBytes);
        final List<CharsetRow> rows = pxl.importCsv()
                .override(option)
                .sheet(CharsetRow.class)
                .fromStream("Cities", inputStream);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getCity()).isEqualTo("서울");
    }

    // ------------------------------------------------------------------
    // Per-sheet charset/delimiter cascade
    // (sheet option > @PxlSheet > workbook option > @PxlWorkbook > built-in default)
    // ------------------------------------------------------------------

    @Test
    public void importCsvStreams_perSheetCharset_decodesEachSheetWithItsOwn() throws Exception {
        // A CSV workbook is one file per sheet, so its sheets need not share an encoding. MixedCharsetWorkbook names
        // MS949 at the workbook level and UTF-8 on "Modern": "Legacy" inherits, "Modern" departs.
        final List<String> csvNames = Arrays.asList("Legacy", "Modern");
        final List<InputStream> csvStreams = Arrays.asList(
                stream(SEOUL_CSV, "MS949"),
                stream(BUSAN_CSV, "UTF-8"));

        final MixedCharsetWorkbook workbook = pxl.importCsv()
                .workbook(MixedCharsetWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getLegacy()).extracting(CharsetRow::getCity).containsExactly("서울");
        assertThat(workbook.getModern()).extracting(CharsetRow::getCity).containsExactly("부산");
    }

    @Test
    public void importCsvStreams_inheritedSheetCharset_isNotTheBuiltInDefault() throws Exception {
        // Feeding "Legacy" UTF-8 bytes proves its inherited MS949 is genuinely applied: read as MS949 they cannot come
        // back as the text they encode. Asserting only "not equal" keeps this off the exact replacement characters.
        final List<String> csvNames = Arrays.asList("Legacy", "Modern");
        final List<InputStream> csvStreams = Arrays.asList(
                stream(SEOUL_CSV, "UTF-8"),
                stream(BUSAN_CSV, "UTF-8"));

        final MixedCharsetWorkbook workbook = pxl.importCsv()
                .workbook(MixedCharsetWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getLegacy().get(0).getCity()).isNotEqualTo("서울");
        // Its sibling is unaffected, so what differs is the one sheet's charset and not a workbook-wide one.
        assertThat(workbook.getModern()).extracting(CharsetRow::getCity).containsExactly("부산");
    }

    @Test
    public void importCsvStreams_sheetCharsetEqualToDefault_stillOverridesWorkbook() throws Exception {
        // "Modern" names UTF-8, which is also the built-in default. Were the annotation's "not specified" marker the
        // effective default rather than a sentinel, this sheet would be indistinguishable from one that says nothing
        // and would silently fall back to the workbook's MS949 - garbling the UTF-8 bytes below.
        final List<String> csvNames = Arrays.asList("Legacy", "Modern");
        final List<InputStream> csvStreams = Arrays.asList(
                stream(SEOUL_CSV, "MS949"),
                stream(SEOUL_CSV, "UTF-8"));

        final MixedCharsetWorkbook workbook = pxl.importCsv()
                .workbook(MixedCharsetWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        // Same text, different bytes, both read correctly - only a per-sheet charset can do that.
        assertThat(workbook.getLegacy()).extracting(CharsetRow::getCity).containsExactly("서울");
        assertThat(workbook.getModern()).extracting(CharsetRow::getCity).containsExactly("서울");
    }

    @Test
    public void importCsvStreams_perSheetDelimiter_splitsEachSheetWithItsOwn() throws Exception {
        // The delimiter counterpart: the workbook names the tab, "Comma" names the comma.
        final List<String> csvNames = Arrays.asList("Tabbed", "Comma");
        final List<InputStream> csvStreams = Arrays.asList(stream(DEPARTMENTS_TSV), stream(DEPARTMENTS_CSV));

        final MixedDelimiterWorkbook workbook = pxl.importCsv()
                .workbook(MixedDelimiterWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getTabbed()).extracting(Department::getCode).containsExactly("ENG");
        assertThat(workbook.getTabbed().get(0).getHeadcount()).isEqualTo(12);
        assertThat(workbook.getComma()).extracting(Department::getCode).containsExactly("ENG", "SAL");
    }

    @Test
    public void importCsvStreams_workbookOptionCharset_losesToSheetAnnotation() throws Exception {
        // The workbook option outranks @PxlWorkbook but not @PxlSheet, so this single override lands on "Legacy"
        // alone - it drops from MS949 to ISO-8859-1, while "Modern" keeps the UTF-8 its own annotation names.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("ISO-8859-1")
                .build();
        final List<String> csvNames = Arrays.asList("Legacy", "Modern");
        final List<InputStream> csvStreams = Arrays.asList(
                stream(CAFE_CSV, "ISO-8859-1"),
                stream(BUSAN_CSV, "UTF-8"));

        final MixedCharsetWorkbook workbook = pxl.importCsv()
                .override(option)
                .workbook(MixedCharsetWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getLegacy()).extracting(CharsetRow::getCity).containsExactly("café");
        assertThat(workbook.getModern()).extracting(CharsetRow::getCity).containsExactly("부산");
    }

    @Test
    public void importCsvStreams_sheetOptionCharset_overridesSheetAnnotation() throws Exception {
        // The runtime sheet option is the top of the cascade: targeted at the "modern" field by name, it displaces
        // the UTF-8 that field's @PxlSheet names. "Legacy" matches no option and keeps its inherited MS949.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .fieldName("modern")
                        .importCsvCharset("ISO-8859-1")
                        .build()))
                .build();
        final List<String> csvNames = Arrays.asList("Legacy", "Modern");
        final List<InputStream> csvStreams = Arrays.asList(
                stream(SEOUL_CSV, "MS949"),
                stream(CAFE_CSV, "ISO-8859-1"));

        final MixedCharsetWorkbook workbook = pxl.importCsv()
                .override(option)
                .workbook(MixedCharsetWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getLegacy()).extracting(CharsetRow::getCity).containsExactly("서울");
        assertThat(workbook.getModern()).extracting(CharsetRow::getCity).containsExactly("café");
    }

    @Test
    public void importCsvStreams_sheetOptionDelimiter_overridesSheetAnnotation() throws Exception {
        // Same precedence for the delimiter: the option pushes the "comma" field back onto the tab.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .fieldName("comma")
                        .importCsvDelimiter('\t')
                        .build()))
                .build();
        final List<String> csvNames = Arrays.asList("Tabbed", "Comma");
        final List<InputStream> csvStreams = Arrays.asList(stream(DEPARTMENTS_TSV), stream(DEPARTMENTS_TSV));

        final MixedDelimiterWorkbook workbook = pxl.importCsv()
                .override(option)
                .workbook(MixedDelimiterWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getTabbed()).extracting(Department::getCode).containsExactly("ENG");
        assertThat(workbook.getComma()).extracting(Department::getCode).containsExactly("ENG");
    }

    @Test
    public void importCsvStreams_workbookNamesNeitherCsvAttribute_usesBuiltInDefaults() throws Exception {
        // Both annotation levels hold the "not specified" sentinel here, so the built-in UTF-8/comma must apply.
        // Taking the sentinel for a usable value instead would hand "" to Charset.forName and fail every sheet.
        final List<String> csvNames = Arrays.asList("Cities", "Departments");
        final List<InputStream> csvStreams = Arrays.asList(stream(SEOUL_CSV, "UTF-8"), stream(DEPARTMENTS_CSV));

        final DefaultCsvWorkbook workbook = pxl.importCsv()
                .workbook(DefaultCsvWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getCities()).extracting(CharsetRow::getCity).containsExactly("서울");
        assertThat(workbook.getDepartments()).extracting(Department::getCode).containsExactly("ENG", "SAL");
    }

    @Test
    public void importCsvStreams_invalidSheetCharset_errorNamesOffendingSheet() throws Exception {
        // With the charset resolved per sheet, a workbook-wide message would leave the caller to guess which of the
        // files is misconfigured, so the message names the sheet the unusable value was resolved for.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .fieldName("modern")
                        .importCsvCharset("NoSuchCharset-1")
                        .build()))
                .build();
        final List<String> csvNames = Arrays.asList("Legacy", "Modern");
        final List<InputStream> csvStreams = Arrays.asList(stream(SEOUL_CSV, "MS949"), stream(BUSAN_CSV, "UTF-8"));

        final PxlArgumentException exception = assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .override(option)
                .workbook(MixedCharsetWorkbook.class)
                .fromStreams(csvNames, csvStreams));

        // Both bundles interpolate the sheet name and the value, so these hold whatever the process locale is.
        assertThat(exception.getMessage())
                .contains("Modern")
                .contains("NoSuchCharset-1")
                .doesNotContain("Legacy");
        assertThat(exception).hasCauseInstanceOf(UnsupportedCharsetException.class);
    }

    @Test
    public void importCsvStream_sheetForm_invalidCharset_errorNamesSheet() throws Exception {
        // The sheet form binds no @PxlSheet field, so its sheet name is the one handed to fromStream.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("NoSuchCharset-1")
                .build();

        final PxlArgumentException exception = assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .override(option)
                .sheet(Employee.class)
                .fromStream("Payroll", stream(EMPLOYEES_CSV)));

        assertThat(exception.getMessage()).contains("Payroll");
    }

    @Test
    public void importCsvStream_sheetForm_wildcardSheetOptionCharset_applies() throws Exception {
        // The sheet form has no field to carry @PxlSheet, so the wildcard sheet option is the only sheet-level
        // route into the cascade - the reason the option had to gain these two fields alongside the annotation.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .importCsvCharset("MS949")
                        .build()))
                .build();

        final List<CharsetRow> rows = pxl.importCsv()
                .override(option)
                .sheet(CharsetRow.class)
                .fromStream("Cities", stream(SEOUL_CSV, "MS949"));

        assertThat(rows).extracting(CharsetRow::getCity).containsExactly("서울");
    }

    @Test
    public void importCsvStream_sheetForm_sheetOptionCharset_overridesWorkbookOption() throws Exception {
        // Both option levels are set; the sheet-level one must win.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("ISO-8859-1")
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .importCsvCharset("MS949")
                        .build()))
                .build();

        final List<CharsetRow> rows = pxl.importCsv()
                .override(option)
                .sheet(CharsetRow.class)
                .fromStream("Cities", stream(SEOUL_CSV, "MS949"));

        assertThat(rows).extracting(CharsetRow::getCity).containsExactly("서울");
    }

    @Test
    public void importCsvStream_blankOptionCharset_fallsBackToBuiltInDefault() throws Exception {
        // Blank means "not specified" at every level, so an explicitly blank option is not an unusable charset name -
        // it simply hands the decision on, ending at UTF-8.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("")
                .build();

        final List<CharsetRow> rows = pxl.importCsv()
                .override(option)
                .sheet(CharsetRow.class)
                .fromStream("Cities", stream(SEOUL_CSV, "UTF-8"));

        assertThat(rows).extracting(CharsetRow::getCity).containsExactly("서울");
    }

    @Test
    public void importCsvStream_nulOptionDelimiter_fallsBackToBuiltInDefault() throws Exception {
        // The delimiter counterpart of the blank charset: NUL hands the decision on and ends at the comma.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvDelimiter('\0')
                .build();

        final List<Department> departments = pxl.importCsv()
                .override(option)
                .sheet(Department.class)
                .fromStream("Departments", stream(DEPARTMENTS_CSV));

        assertThat(departments).extracting(Department::getCode).containsExactly("ENG", "SAL");
    }

    @Test
    public void importCsvStream_blankSheetOptionCharset_fallsBackToWorkbookOption() throws Exception {
        // A blank at the top of the cascade must not shadow the level below it: the workbook option still decides.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importCsvCharset("MS949")
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .importCsvCharset("")
                        .build()))
                .build();

        final List<CharsetRow> rows = pxl.importCsv()
                .override(option)
                .sheet(CharsetRow.class)
                .fromStream("Cities", stream(SEOUL_CSV, "MS949"));

        assertThat(rows).extracting(CharsetRow::getCity).containsExactly("서울");
    }

    @Test
    public void importCsvFiles_perSheetCharset_decodesEachFileWithItsOwn(@TempDir final Path tempDir) throws Exception {
        // The file-name form of the same thing: each file carries its own encoding, and the name (extension
        // removed) still selects the sheet whose charset decodes it.
        final Path legacyPath = tempDir.resolve("Legacy.csv");
        final Path modernPath = tempDir.resolve("Modern.csv");
        Files.write(legacyPath, SEOUL_CSV.getBytes(Charset.forName("MS949")));
        Files.write(modernPath, BUSAN_CSV.getBytes(StandardCharsets.UTF_8));

        final MixedCharsetWorkbook workbook = pxl.importCsv()
                .workbook(MixedCharsetWorkbook.class)
                .fromFiles(Arrays.asList(legacyPath.toFile(), modernPath.toFile()));

        assertThat(workbook.getLegacy()).extracting(CharsetRow::getCity).containsExactly("서울");
        assertThat(workbook.getModern()).extracting(CharsetRow::getCity).containsExactly("부산");
    }

    @Test
    public void importCsvStreams_blankSheetOptionCharset_fallsBackToSheetAnnotation() throws Exception {
        // A blank sheet option hands the decision to the level directly below it, which is the sheet's own
        // annotation - not all the way down to the workbook. "Modern" must therefore stay UTF-8, not become MS949.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .fieldName("modern")
                        .importCsvCharset("")
                        .build()))
                .build();
        final List<String> csvNames = Arrays.asList("Legacy", "Modern");
        final List<InputStream> csvStreams = Arrays.asList(
                stream(SEOUL_CSV, "MS949"),
                stream(BUSAN_CSV, "UTF-8"));

        final MixedCharsetWorkbook workbook = pxl.importCsv()
                .override(option)
                .workbook(MixedCharsetWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getLegacy()).extracting(CharsetRow::getCity).containsExactly("서울");
        assertThat(workbook.getModern()).extracting(CharsetRow::getCity).containsExactly("부산");
    }

    @Test
    public void importCsvStreams_nulSheetOptionDelimiter_fallsBackToSheetAnnotation() throws Exception {
        // The delimiter counterpart: NUL on the "comma" field leaves its @PxlSheet comma standing rather than
        // dropping through to the workbook's tab.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .fieldName("comma")
                        .importCsvDelimiter('\0')
                        .build()))
                .build();
        final List<String> csvNames = Arrays.asList("Tabbed", "Comma");
        final List<InputStream> csvStreams = Arrays.asList(stream(DEPARTMENTS_TSV), stream(DEPARTMENTS_CSV));

        final MixedDelimiterWorkbook workbook = pxl.importCsv()
                .override(option)
                .workbook(MixedDelimiterWorkbook.class)
                .fromStreams(csvNames, csvStreams);

        assertThat(workbook.getTabbed()).extracting(Department::getCode).containsExactly("ENG");
        assertThat(workbook.getComma()).extracting(Department::getCode).containsExactly("ENG", "SAL");
    }

    @Test
    public void importCsvStream_sheetForm_wildcardSheetOptionDelimiter_applies() throws Exception {
        // The sheet form's delimiter route, the counterpart of the wildcard charset above: with no @PxlSheet
        // field to read, the ad-hoc sheet meta resolves the delimiter from the wildcard option alone.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .importCsvDelimiter('\t')
                        .build()))
                .build();

        final List<Department> departments = pxl.importCsv()
                .override(option)
                .sheet(Department.class)
                .fromStream("Departments", stream(DEPARTMENTS_TSV));

        assertThat(departments).extracting(Department::getCode).containsExactly("ENG");
        assertThat(departments.get(0).getHeadcount()).isEqualTo(12);
    }

    @Test
    public void importCsvStream_invalidDelimiterOnSheetAnnotation_errorNamesSheet() throws Exception {
        // The same workbook the Excel path reads without complaint (see PxlExcelImportTests): read as CSV, the
        // quote character @PxlSheet declares as this sheet's delimiter is rejected, and the sheet is named.
        final PxlArgumentException exception = assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .workbook(IgnoredCsvAttrsWorkbook.class)
                .fromStream("Employees", stream(EMPLOYEES_CSV)));

        assertThat(exception.getMessage()).contains("Employees");
        assertThat(exception).hasMessageContaining("importCsvDelimiter");
        assertThat(exception).hasCauseInstanceOf(IllegalArgumentException.class);
    }

    // ------------------------------------------------------------------
    // Value conversion: CSV shares the codecs with Excel, but always through the string path
    // ------------------------------------------------------------------

    @Test
    public void importCsvStream_numberCustomPatternTrailingGarbage_throws() throws Exception {
        // Every CSV value arrives as a string, so a numeric custom pattern is the parser for every row here - the
        // branch Excel only reaches through a STRING cell. The pattern has to consume the value in full, so "123abc"
        // is rejected instead of binding 123 (issue M3 fix).
        final PxlCellCodecException exception = assertThrows(PxlCellCodecException.class, () -> pxl.importCsv()
                .sheet(NumberPatternRow.class)
                .fromStream("Numbers", stream("WrapInt\n123abc\n")));

        assertThat(exception.getCause()).hasMessageContaining("123abc");
    }

    @Test
    public void importCsvStream_numberCustomPatternGroupedValue_binds() throws Exception {
        // The counterpart: a value the pattern reads end to end still binds, so the stricter parse costs nothing here.
        final List<NumberPatternRow> rows = pxl.importCsv()
                .sheet(NumberPatternRow.class)
                .fromStream("Numbers", stream("WrapInt\n\"1,234\"\n"));

        assertThat(rows.get(0).getWrapInt()).isEqualTo(1234);
    }

    // ------------------------------------------------------------------
    // Format limits (PxlFileFormat.CSV): sheets/columns per sheet
    // ------------------------------------------------------------------

    @Test
    public void importCsvStream_columnCountOverLimit_throws() throws Exception {
        // The column cap is counted off the header record rather than off the bound columns, so one column past
        // it fails even though the row class binds a single one.
        final StringBuilder header = new StringBuilder("City");
        final StringBuilder data = new StringBuilder("Seoul");
        for (int columnIndex = 1; columnIndex <= PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_COLUMNS; columnIndex++) {
            header.append(",Extra").append(columnIndex);
            data.append(",x");
        }
        final String wideCsv = header + "\n" + data + "\n";

        final PxlDataException exception = assertThrows(PxlDataException.class, () -> pxl.importCsv()
                .sheet(CharsetRow.class)
                .fromStream("Cities", stream(wideCsv)));

        assertThat(exception).hasMessageContaining("Cities");
        assertThat(exception.getMessage()).contains(String.valueOf(PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_COLUMNS));
    }

    @Test
    public void importCsvStreams_sheetCountOverLimit_throws() throws Exception {
        // One CSV is one sheet, so the sheet limit bounds how many sources may be handed over at once. It is
        // checked before any of them is matched to a @PxlSheet field, so the names need not correspond to any.
        final int sourceCount = PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_SHEETS + 1;
        final List<String> csvNames = new ArrayList<>();
        final List<InputStream> csvStreams = new ArrayList<>();
        for (int sourceIndex = 0; sourceIndex < sourceCount; sourceIndex++) {
            csvNames.add("Sheet" + sourceIndex);
            csvStreams.add(stream(DEPARTMENTS_CSV));
        }

        final PxlDataException exception = assertThrows(PxlDataException.class, () -> pxl.importCsv()
                .workbook(DefaultCsvWorkbook.class)
                .fromStreams(csvNames, csvStreams));

        assertThat(exception.getMessage()).contains(String.valueOf(PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_SHEETS));
    }

    // ------------------------------------------------------------------
    // Explicit (non-default) header/data row and column indices
    // ------------------------------------------------------------------

    @Test
    public void importCsvStream_explicitRowAndColumnIndices_selectsRegion() throws Exception {
        // Non-default 1-based indices exercise the importer's 1-based -> 0-based bound calculations
        // (the "else" branches a default-positioned CSV skips), for both rows and columns.
        final String csv =
                "junk0,junk1,junk2,junk3\n" +   // 0-based row 0: before the header
                        "pad,Name,Age,extra\n" +        // 0-based row 1: header (only cols 1..2 are used)
                        "pad,Alice,30,extra\n" +        // 0-based row 2: first data row
                        "pad,Bob,42,extra\n" +          // 0-based row 3: last data row
                        "pad,Carol,99,extra\n";         // 0-based row 4: beyond the last data row -> excluded

        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(2)          // 1-based -> 0-based 1
                .importFirstDataRowIndex(3)       // 1-based -> 0-based 2
                .importLastDataRowIndex(4)        // 1-based -> 0-based 3 (inclusive)
                .importFirstDataColumnIndex(2)    // 1-based -> 0-based 1
                .importLastDataColumnIndex(3)     // 1-based -> 0-based 2 (inclusive)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importDataValidation(false)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<Employee> employees = pxl.importCsv()
                .override(option)
                .sheet(Employee.class)
                .fromStream("Employees", stream(csv));

        // only the Name/Age columns (0-based 1..2) and the rows 0-based 2..3 are read; Carol is excluded
        assertThat(employees).extracting(Employee::getName).containsExactly("Alice", "Bob");
        assertThat(employees.get(0).getAge()).isEqualTo(30);
        assertThat(employees.get(1).getAge()).isEqualTo(42);
    }

    // ------------------------------------------------------------------
    // @PxlRowIndex field-type injection (Long / Short / Byte / Number, and an unsupported type)
    // ------------------------------------------------------------------

    @Test
    public void importCsv_rowIndexFieldTypes_injectedPerType() throws Exception {
        // The 1-based row number is injected into the @PxlRowIndex field, one branch per numeric field type.
        final String csv = "Name\nAlice\n";
        assertThat(pxl.importCsv()
                .sheet(LongRowIndexRow.class)
                .fromStream("S", stream(csv)).get(0).getRowIndex()).isEqualTo(2L);
        assertThat(pxl.importCsv()
                .sheet(ShortRowIndexRow.class)
                .fromStream("S", stream(csv)).get(0).getRowIndex()).isEqualTo((short) 2);
        assertThat(pxl.importCsv()
                .sheet(ByteRowIndexRow.class)
                .fromStream("S", stream(csv)).get(0).getRowIndex()).isEqualTo((byte) 2);
        assertThat(pxl.importCsv()
                .sheet(NumberRowIndexRow.class)
                .fromStream("S", stream(csv)).get(0).getRowIndex().intValue()).isEqualTo(2);
        // an unsupported @PxlRowIndex type (String) is rejected when the int index cannot be cast to it
        assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .sheet(BadRowIndexRow.class)
                .fromStream("S", stream(csv)));
    }

    @Test
    public void importCsv_rowIndex_multipleRowsShiftedLayout_injectsOneBasedRowNumber() throws Exception {
        // The @PxlRowIndex value is the 1-based row number of each imported row, not a sequential counter:
        // with the header shifted to 0-based record 2, the three data records at 0-based records 3/4/5 receive
        // exactly 4/5/6 (proving both per-row distinctness and the row-number semantics).
        final String csv = "junk\njunk\nName\nAlice\nBob\nCarol\n";
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .importHeaderRowIndex(3)          // 1-based -> 0-based 2
                        .importFirstDataRowIndex(4)       // 1-based -> 0-based 3
                        .build()))
                .build();

        final List<LongRowIndexRow> rows = pxl.importCsv()
                .override(option)
                .sheet(LongRowIndexRow.class)
                .fromStream("S", stream(csv));

        assertThat(rows).extracting(LongRowIndexRow::getName).containsExactly("Alice", "Bob", "Carol");
        assertThat(rows).extracting(LongRowIndexRow::getRowIndex).containsExactly(4L, 5L, 6L);
    }

    // ------------------------------------------------------------------
    // Header/column resolution errors
    // ------------------------------------------------------------------

    @Test
    public void importCsv_noMatchingColumn_throws() {
        // No header cell matches any column name -> "no header column" error.
        final String csv = "Foo\n1\n";   // "Foo" != "Name"
        assertThrows(PxlDataException.class, () -> pxl.importCsv()
                .sheet(LongRowIndexRow.class)
                .fromStream("S", stream(csv)));
    }

    @Test
    public void importCsv_headerRowIndexBeyondData_throws() {
        // A header row index past the end of the records -> "no header row" error.
        final String csv = "Name\nAlice\n";
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder().importHeaderRowIndex(10).build()))
                .build();
        assertThrows(PxlDataException.class,
                () -> pxl.importCsv()
                        .override(option)
                        .sheet(LongRowIndexRow.class)
                        .fromStream("S", stream(csv)));
    }

    @Test
    public void importCsv_duplicateHeaderColumn_throws() {
        // The same column name appears twice and both match one column -> duplicate-column error.
        final String csv = "Name,Name\nAlice,Bob\n";
        assertThrows(PxlDataException.class, () -> pxl.importCsv()
                .sheet(LongRowIndexRow.class)
                .fromStream("S", stream(csv)));
    }

    @Test
    public void importCsvFile_workbookForm_decomposedFileName_matchesComposedSheetName(@TempDir final Path tempDir) throws Exception {
        // macOS file systems hand back decomposed (NFD) file names, so the derived sheet name is normalized to NFC -
        // otherwise it would not equal the composed name written in @PxlSheet and the sheet would not be found.
        // Creating the file with an NFD name reproduces that on any OS, since File.getName() returns the name as given.
        final String composed = "직원";                                   // one of the @PxlSheet aliases of AliasSheetWorkbook
        final String decomposed = Normalizer.normalize(composed, Normalizer.Form.NFD);
        assertThat(decomposed).isNotEqualTo(composed);

        final File employeesCsv = writeCsv(tempDir, decomposed + ".csv", EMPLOYEES_CSV);

        final AliasSheetWorkbook workbook = pxl.importCsv()
                .workbook(AliasSheetWorkbook.class)
                .fromFile(employeesCsv);

        assertEmployees(workbook.getData());
    }

    @Test
    public void importCsvFile_workbookForm_differentCaseFileName_matchesSheetName(@TempDir final Path tempDir) throws Exception {
        // A file name carries whatever casing the file system holds - Windows does not distinguish it at all - so the
        // derived sheet name is matched against @PxlSheet ignoring case: "EMPLOYEES.csv" still selects "Employees".
        final File employeesCsv = writeCsv(tempDir, "EMPLOYEES.csv", EMPLOYEES_CSV);

        final CompanyWorkbook workbook = pxl.importCsv()
                .workbook(CompanyWorkbook.class)
                .fromFile(employeesCsv);

        assertEmployees(workbook.getEmployees());
        assertThat(workbook.getDepartments()).isNullOrEmpty();
    }

    @Test
    public void importCsvStream_workbookForm_differentCaseCsvName_matchesSheetName() throws Exception {
        // The explicitly given csvName is matched the same way as a file-derived one.
        final CompanyWorkbook workbook = pxl.importCsv()
                .workbook(CompanyWorkbook.class)
                .fromStream("employees", stream(EMPLOYEES_CSV));

        assertEmployees(workbook.getEmployees());
    }

    @Test
    public void importCsv_shortAndBlankRows_skipped() throws Exception {
        // A row shorter than the mapped column index leaves that field null; a row whose only mapped column is
        // blank is treated as an ignorable row and skipped.
        final String csv = "Ignore,Name\nfoo,Alice\nbar\n";   // "bar" has no Name cell (col 1 missing)
        final List<LongRowIndexRow> rows = pxl.importCsv()
                .sheet(LongRowIndexRow.class)
                .fromStream("S", stream(csv));
        assertThat(rows).extracting(LongRowIndexRow::getName).containsExactly("Alice");
    }
}
