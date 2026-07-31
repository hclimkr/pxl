package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.builder.PxlCsvImportBuilder;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
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
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;
import java.util.Set;

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

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    private static InputStream stream(final String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.US_ASCII));
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
