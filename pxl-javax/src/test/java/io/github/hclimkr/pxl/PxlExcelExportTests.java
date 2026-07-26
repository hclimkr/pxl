package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.option.*;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.time.*;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Excel export path tests — masking/trimming/grouping/password/file format (SXSSF, HSSF), literal formulas, column inheritance, sheet name normalization, lastColumnIndex boundary.
 */
public class PxlExcelExportTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    // Captures the current test method name to match the export file name.
    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

    // Exports the given rows into a single sheet in a real file, then imports it back. (file name = test method name)
    private <T> List<T> roundTripSheet(final String sheetName, final List<T> rows, final Class<T> rowClass) throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(sheetName, rows, rowClass)
                .override(noValidationOption())
                .toFile(excelFile);
        return pxl.importExcel()
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromFile(excelFile);
    }

    private static List<Employee> twoEmployees() {
        return Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"));
    }

    // The actual (normalized) name of the first sheet
    private static String firstSheetName(final byte[] bytes) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            return workbook.getSheetAt(0).getSheetName();
        }
    }

    // ------------------------------------------------------------------
    // Multiple sheets — a different rowClass per sheet
    // ------------------------------------------------------------------

    @Test
    public void exportMultiSheet_perSheetRowClass_appliesEachClass() throws Exception {
        final List<Employee> employees = twoEmployees();
        final List<AllTypesRow> allTypes = Arrays.asList(Fixtures.sampleAllTypesRow());

        final File excelFile = TestPaths.exportFile(testInfo);
        // Call sheet() multiple times to specify a different rowClass per sheet.
        pxl.exportExcel()
                .sheet("Employees", employees, Employee.class)
                .sheet("AllTypes", allTypes, AllTypesRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final List<Employee> importedEmployees = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Employees"))
                .fromFile(excelFile);
        final List<AllTypesRow> importedAllTypes = pxl.importExcel()
                .sheet(AllTypesRow.class, Arrays.asList("AllTypes"))
                .fromFile(excelFile);

        // Verify each sheet is bound with its own rowClass — different classes are applied per sheet.
        assertThat(importedEmployees).extracting(Employee::getName).containsExactly("Alice", "Bob");
        assertThat(importedAllTypes).hasSize(1);
        assertThat(importedAllTypes.get(0).getText()).isEqualTo("Hello, PXL");
    }

    @Test
    public void exportMultiSheet_threeSheets_preservesCallOrderAndRows() throws Exception {
        final List<Employee> engineering = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"));
        final List<Employee> sales = Arrays.asList(
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"),
                Fixtures.employee("Carol", 35, "68000", true, LocalDate.of(2019, 3, 20), Grade.A, "Sales"));
        final List<Department> departments = Arrays.asList(
                Fixtures.department("ENG", "Engineering", 12),
                Fixtures.department("SAL", "Sales", 7));

        final File excelFile = TestPaths.exportFile(testInfo);
        // Three sheet() calls -> three sheets; the same rowClass may repeat and the call order is the sheet order.
        pxl.exportExcel()
                .sheet("Engineering", engineering, Employee.class)
                .sheet("Sales", sales, Employee.class)
                .sheet("Departments", departments, Department.class)
                .override(noValidationOption())
                .toFile(excelFile);

        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            assertThat(workbook.getNumberOfSheets()).isEqualTo(3);
            assertThat(workbook.getSheetName(0)).isEqualTo("Engineering");
            assertThat(workbook.getSheetName(1)).isEqualTo("Sales");
            assertThat(workbook.getSheetName(2)).isEqualTo("Departments");
        }

        // Rows must not leak across sheets: each sheet holds only the rows passed with its own sheet() call.
        final List<Employee> importedEngineering = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Engineering"))
                .fromFile(excelFile);
        final List<Employee> importedSales = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Sales"))
                .fromFile(excelFile);
        final List<Department> importedDepartments = pxl.importExcel()
                .sheet(Department.class, Arrays.asList("Departments"))
                .fromFile(excelFile);

        assertThat(importedEngineering).extracting(Employee::getName).containsExactly("Alice");
        assertThat(importedSales).extracting(Employee::getName).containsExactly("Bob", "Carol");
        assertThat(importedDepartments).extracting(Department::getCode).containsExactly("ENG", "SAL");
        assertThat(importedDepartments.get(0).getHeadcount()).isEqualTo(12);
    }

    // ------------------------------------------------------------------
    // Masking (exportMasking)
    // ------------------------------------------------------------------

    @Test
    public void exportMasking_digits_replaced() throws Exception {
        final MaskingRow row = new MaskingRow();
        row.setSecret("ID12345");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Masking", Arrays.asList(row), MaskingRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final List<MaskingRow> rows = pxl.importExcel()
                .sheet(MaskingRow.class, Arrays.asList("Masking"))
                .fromFile(excelFile);

        assertThat(rows).hasSize(1);
        // All 5 digits are replaced with '*'.
        assertThat(rows.get(0).getSecret()).isEqualTo("ID*****");
    }

    // ------------------------------------------------------------------
    // Export trim (exportTrim)
    // ------------------------------------------------------------------

    @Test
    public void exportTrim_whitespace_trimmed() throws Exception {
        final TrimRow row = new TrimRow();
        row.setPadded("  spaced  ");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Trim", Arrays.asList(row), TrimRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final List<TrimRow> rows = pxl.importExcel()
                .sheet(TrimRow.class, Arrays.asList("Trim"))
                .fromFile(excelFile);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getPadded()).isEqualTo("spaced");
    }

    // ------------------------------------------------------------------
    // Grouping (exportGroupingFieldName) - sheets are split by department value.
    // ------------------------------------------------------------------

    @Test
    public void exportGrouping_splitsIntoSheets() throws Exception {
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"),
                Fixtures.employee("Carol", 35, "68000", true, LocalDate.of(2019, 3, 20), Grade.A, "Engineering")));

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toFile(excelFile);

        // Group sheet names follow the "<sheet name> - <group value>" format.
        try (Workbook poiWorkbook = WorkbookFactory.create(excelFile)) {
            assertThat(poiWorkbook.getSheet("Employees - Engineering")).as("Engineering group sheet is missing.").isNotNull();
            assertThat(poiWorkbook.getSheet("Employees - Sales")).as("Sales group sheet is missing.").isNotNull();
        }

        final List<Employee> engineering = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Employees - Engineering"))
                .fromFile(excelFile);
        final List<Employee> sales = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Employees - Sales"))
                .fromFile(excelFile);

        assertThat(engineering).extracting(Employee::getName).containsExactly("Alice", "Carol");
        assertThat(sales).extracting(Employee::getName).containsExactly("Bob");
    }

    // ------------------------------------------------------------------
    // Password (exportPassword / importPassword) round trip
    // ------------------------------------------------------------------

    @Test
    public void exportPassword_encrypted_roundTrips() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(exportOption)
                .toFile(excelFile);

        final PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
                .importPassword("secret")
                .build();
        final List<Employee> people = pxl.importExcel()
                .override(importOption)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
    }

    // ------------------------------------------------------------------
    // File format: SXSSF (streaming write, result is xlsx)
    // ------------------------------------------------------------------

    @Test
    public void exportSxssf_streaming_roundTrips() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.SXSSF)
                .exportDataValidation(false)
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(option)
                .toFile(excelFile);

        final List<Employee> people = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
    }

    // ------------------------------------------------------------------
    // File format: HSSF (xls) round trip
    // ------------------------------------------------------------------

    @Test
    public void exportHssf_xls_roundTrips() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.HSSF)
                .exportDataValidation(false)
                .build();

        final File xlsFile = TestPaths.exportFile(testInfo, ".xls");
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(option)
                .toFile(xlsFile);

        final List<Employee> people = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(xlsFile);

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
        assertThat(people.get(0).getGrade()).isEqualTo(Grade.A);
    }

    // ------------------------------------------------------------------
    // File format x password: SXSSF/HSSF encrypted round trip (writeToStream uses agile encryption for both XSSF and SXSSF, Biff8 for HSSF)
    // ------------------------------------------------------------------

    @Test
    public void exportSxssf_encrypted_roundTrips() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.SXSSF)
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(exportOption)
                .toFile(excelFile);

        final PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
                .importPassword("secret")
                .build();
        final List<Employee> people = pxl.importExcel()
                .override(importOption)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
        assertThat(people.get(0).getGrade()).isEqualTo(Grade.A);
    }

    @Test
    public void exportHssf_encrypted_roundTrips() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.HSSF)
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();
        final File xlsFile = TestPaths.exportFile(testInfo, ".xls");
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(exportOption)
                .toFile(xlsFile);

        final PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
                .importPassword("secret")
                .build();
        final List<Employee> people = pxl.importExcel()
                .override(importOption)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(xlsFile);

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");

        // The Biff8 thread-local key is cleaned up in finally, so a subsequent unencrypted HSSF export is not contaminated.
        final PxlExportWorkbookOption plainOption = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.HSSF)
                .exportDataValidation(false)
                .build();
        final File plainXls = TestPaths.exportFile("hssf-plain-after-encrypted.xls");
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(plainOption)
                .toFile(plainXls);
        final List<Employee> plainPeople = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(plainXls);
        assertThat(plainPeople).hasSize(1);
    }

    // ------------------------------------------------------------------
    // Resource ownership: toStream does not close the OutputStream passed by the caller (only flushes)
    // ------------------------------------------------------------------

    @Test
    public void exportToStream_doesNotCloseCallerStream() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final boolean[] closed = {false};
        final ByteArrayOutputStream tracking = new ByteArrayOutputStream() {
            @Override
            public void close() {
                closed[0] = true;
            }
        };

        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(noValidationOption())
                .toStream(tracking);

        assertThat(closed[0]).as("toStream must not close the caller's stream").isFalse();
        assertThat(tracking.size()).isGreaterThan(0);
    }

    // ------------------------------------------------------------------
    // exportStringAsFormula=false: "=..." strings are not evaluated as formulas but preserved as literals
    // ------------------------------------------------------------------

    @Test
    public void exportLiteralFormula_notEvaluated() throws Exception {
        final LiteralRow row = new LiteralRow();
        row.setExpr("=1+2");

        final List<LiteralRow> out = roundTripSheet("Literal", Arrays.asList(row), LiteralRow.class);

        assertThat(out).hasSize(1);
        // If evaluated as a formula it would become "3", but as a literal it must stay "=1+2".
        assertThat(out.get(0).getExpr()).isEqualTo("=1+2");
    }

    // ------------------------------------------------------------------
    // Column inheritance: @PxlColumn fields from the superclass are also bound
    // ------------------------------------------------------------------

    @Test
    public void exportInheritedColumns_roundTrips() throws Exception {
        final DerivedRow row = new DerivedRow();
        row.setId(7);                 // inherited field
        row.setBaseName("base");      // inherited field
        row.setExtra("own");          // own field

        final List<DerivedRow> out = roundTripSheet("Derived", Arrays.asList(row), DerivedRow.class);

        assertThat(out).hasSize(1);
        assertThat(out.get(0).getId()).isEqualTo(7);
        assertThat(out.get(0).getBaseName()).isEqualTo("base");
        assertThat(out.get(0).getExtra()).isEqualTo("own");
    }

    // Separately verify that a superclass-only field actually appears in the header
    @Test
    public void exportInheritedColumns_headerIncludesThem() throws Exception {
        final DerivedRow row = new DerivedRow();
        row.setId(1);
        row.setBaseName("base");
        row.setExtra("own");

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .sheet("Derived", Arrays.asList(row), DerivedRow.class)
                .override(noValidationOption())
                .toStream(outputStream);

        try (org.apache.poi.ss.usermodel.Workbook workbook =
                     org.apache.poi.ss.usermodel.WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            final org.apache.poi.ss.usermodel.Row header = workbook.getSheet("Derived").getRow(0);
            final java.util.Set<String> headers = new java.util.HashSet<>();
            for (final org.apache.poi.ss.usermodel.Cell cell : header) {
                headers.add(cell.getStringCellValue());
            }
            assertThat(headers).contains("Id", "BaseName", "Extra");
        }
    }

    // ------------------------------------------------------------------
    // All sheets exportEnabled=false: fails because there is no sheet to export
    // ------------------------------------------------------------------

    @Test
    public void exportAllSheetsDisabled_throws() {
        final DisabledSheetWorkbook workbook = new DisabledSheetWorkbook();
        workbook.setWorkbookName("NoSheets");
        workbook.setRows(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering")));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlDataException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // Empty password -> round trips without encryption (importable without a password)
    // ------------------------------------------------------------------

    @Test
    public void exportEmptyPassword_noEncryption_roundTrips() throws Exception {
        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportPassword("")     // empty password = no encryption
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("People", twoEmployees(), Employee.class)
                .override(exportOption)
                .toFile(excelFile);

        // Must be importable without any password option
        final List<Employee> people = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);
        assertThat(people).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Import with password + stream reader combination
    // ------------------------------------------------------------------

    @Test
    public void exportPassword_withStreamReader_roundTrips() throws Exception {
        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("People", twoEmployees(), Employee.class)
                .override(exportOption)
                .toFile(excelFile);

        // Import with the correct password + stream reader (header/data row indices specified)
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
                .importPassword("secret")
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<Employee> people = pxl.importExcel()
                .override(importOption)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);
        assertThat(people).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Sheet name over 31 chars -> truncated to 31 chars
    // ------------------------------------------------------------------

    @Test
    public void exportSheetName_over31Chars_truncated() throws Exception {
        final String longName = "VeryLongSheetNameThatExceeds31Characters";   // 40 chars
        assertThat(longName.length()).isGreaterThan(31);

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(longName, twoEmployees(), Employee.class)
                .override(noValidationOption())
                .toFile(excelFile);
        final byte[] bytes = java.nio.file.Files.readAllBytes(excelFile.toPath());

        // The actual sheet name is truncated to 31 chars
        final String actualName = firstSheetName(bytes);
        assertThat(actualName).hasSize(31);
        assertThat(longName).startsWith(actualName);

        // Importing with the truncated actual name round trips correctly
        final List<Employee> people = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList(actualName))
                .fromFile(excelFile);
        assertThat(people).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Invalid characters in sheet name -> replaced with spaces
    // ------------------------------------------------------------------

    @Test
    public void exportSheetName_invalidChars_sanitized() throws Exception {
        // Invalid characters in Excel sheet names: / \ ? * [ ] :
        final String badName = "Bad/Name:With*Chars";

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(badName, twoEmployees(), Employee.class)
                .override(noValidationOption())
                .toFile(excelFile);
        final byte[] bytes = java.nio.file.Files.readAllBytes(excelFile.toPath());

        final String actualName = firstSheetName(bytes);
        // No invalid characters must remain
        assertThat(actualName).doesNotContain("/", "\\", "?", "*", "[", "]", ":");
        // Invalid characters are replaced with spaces
        assertThat(actualName).isEqualTo("Bad Name With Chars");

        // Importing with the normalized name round trips correctly
        final List<Employee> people = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList(actualName))
                .fromFile(excelFile);
        assertThat(people).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Exception when the export column range is smaller than the column count
    // ------------------------------------------------------------------

    @Test
    public void exportLastColumnIndex_tooSmall_throws() {
        final ExportColBoundWorkbook workbook = new ExportColBoundWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering")));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlDataException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .override(PxlExportWorkbookOption.builder().exportDataValidation(false).build())
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // Date/time numeric export (when neither pattern nor masking is set, written as numeric date cells rather than strings)
    // ------------------------------------------------------------------

    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    // Finds the cell in the first data row (1) by the header name in the header row (0).
    private static Cell firstDataCell(final Workbook workbook, final String sheetName, final String header) {
        final Sheet sheet = workbook.getSheet(sheetName);
        final Row headerRow = sheet.getRow(0);
        int col = -1;
        for (final Cell cell : headerRow) {
            if (header.equals(cell.getStringCellValue())) {
                col = cell.getColumnIndex();
                break;
            }
        }
        assertThat(col).as("column for header '" + header + "' not found").isGreaterThanOrEqualTo(0);
        return sheet.getRow(1).getCell(col);
    }

    @Test
    public void exportExcel_dateTimeColumnsWithoutPatternOrMasking_writtenAsNumericDateCells() throws Exception {
        // A dedicated fixture whose date/time columns carry no pattern (AllTypesRow now pins explicit patterns,
        // which would force string cells); without a pattern each column is written as a numeric Excel-date cell.
        final ZoneId zone = ZoneId.systemDefault();
        final LocalDateTime baseDateTime = LocalDateTime.of(2023, 6, 15, 10, 30, 45);

        final DateTimeNumericRow row = new DateTimeNumericRow();
        row.setJavaDate(Date.from(baseDateTime.atZone(zone).toInstant()));
        row.setLocalDate(LocalDate.of(2023, 6, 15));
        row.setLocalTime(LocalTime.of(10, 30, 45));
        row.setLocalDateTime(baseDateTime);
        row.setZonedDateTime(baseDateTime.atZone(zone));
        row.setOffsetTime(LocalTime.of(10, 30, 45).atOffset(OffsetTime.now(zone).getOffset()));
        row.setOffsetDateTime(baseDateTime.atZone(zone).toOffsetDateTime());

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("DateTimes", Arrays.asList(row), DateTimeNumericRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final String[] dateHeaders = {"JavaDate", "LocalDate", "LocalTime", "LocalDateTime", "ZonedDateTime", "OffsetTime", "OffsetDateTime"};
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            for (final String header : dateHeaders) {
                final Cell cell = firstDataCell(workbook, "DateTimes", header);
                assertThat(cell.getCellType()).as(header + " cell type").isEqualTo(CellType.NUMERIC);
                assertThat(DateUtil.isCellDateFormatted(cell)).as(header + " date formatting").isTrue();
            }
            // Verify the per-type display format (Excel builtin number format) is applied (checks the format code itself, locale-independently)
            assertThat(firstDataCell(workbook, "DateTimes", "JavaDate").getCellStyle().getDataFormatString()).isEqualTo("m/d/yy");
            assertThat(firstDataCell(workbook, "DateTimes", "LocalDate").getCellStyle().getDataFormatString()).isEqualTo("m/d/yy");
            assertThat(firstDataCell(workbook, "DateTimes", "LocalTime").getCellStyle().getDataFormatString()).isEqualTo("h:mm:ss");
            assertThat(firstDataCell(workbook, "DateTimes", "LocalDateTime").getCellStyle().getDataFormatString()).isEqualTo("m/d/yy h:mm");
        }

        // Even when written as numeric, the values are preserved on round trip (all values are second-granular).
        final List<DateTimeNumericRow> imported = pxl.importExcel()
                .sheet(DateTimeNumericRow.class, Arrays.asList("DateTimes"))
                .fromFile(excelFile);
        assertThat(imported).hasSize(1);
        final DateTimeNumericRow result = imported.get(0);
        assertThat(result.getJavaDate()).isEqualTo(row.getJavaDate());
        assertThat(result.getLocalDate()).isEqualTo(row.getLocalDate());
        assertThat(result.getLocalTime()).isEqualTo(row.getLocalTime());
        assertThat(result.getLocalDateTime()).isEqualTo(row.getLocalDateTime());
        assertThat(result.getZonedDateTime()).isEqualTo(row.getZonedDateTime());
        assertThat(result.getOffsetTime()).isEqualTo(row.getOffsetTime());
        assertThat(result.getOffsetDateTime()).isEqualTo(row.getOffsetDateTime());
    }

    @Test
    public void exportExcel_dateColumnWithMasking_staysStringCell() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        // A date column with masking cannot be represented as numeric, so it stays a (masked) string.
        final PxlExportColumnOption maskHireDate = PxlExportColumnOption.builder()
                .fieldName("hireDate")
                .exportMasking("\\d")
                .build();
        final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                .exportColumnOptions(Arrays.asList(maskHireDate))
                .build();
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(sheetOption))
                .build();

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(option)
                .toStream(outputStream);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            final Cell cell = firstDataCell(workbook, "People", "HireDate");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(cell.getStringCellValue()).isEqualTo("****-**-**");   // all digits of 2020-01-15 are masked
        }
    }

    @Test
    public void exportExcel_dateColumnWithExportPattern_staysStringCell() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        // A date column with exportPattern must keep its string representation, so it is not converted to numeric.
        final PxlExportColumnOption patternHireDate = PxlExportColumnOption.builder()
                .fieldName("hireDate")
                .exportPattern("yyyy/MM/dd")
                .build();
        final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                .exportColumnOptions(Arrays.asList(patternHireDate))
                .build();
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(sheetOption))
                .build();

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice), Employee.class)
                .override(option)
                .toStream(outputStream);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            final Cell cell = firstDataCell(workbook, "People", "HireDate");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(cell.getStringCellValue()).isEqualTo("2020/01/15");
        }
    }

    // ------------------------------------------------------------------
    // exportStringAsFormula: isolate and verify export-time formula computation (caching) with raw POI
    // (the import path re-evaluates formulas, so a round-trip test cannot prove export-time computation)
    // ------------------------------------------------------------------

    // Non-streaming (XSSF): export's evaluateAll() computes the formula and writes the cached result to the file.
    @Test
    public void exportFormula_nonStreaming_cachesComputedResult() throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=2+3");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Formula", Arrays.asList(row), FormulaRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        // Open with raw POI but do not evaluate here — getNumericCellValue() returns the cached value left by export.
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            final Cell cell = firstDataCell(workbook, "Formula", "Formula");
            assertThat(cell.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("2+3");
            // If export had not computed it, the cached value would be 0.0 — 5.0 means it was computed at export time.
            assertThat(cell.getNumericCellValue()).isEqualTo(5.0);
        }
    }

    // A cell-reference formula (=A2*B2) is also computed at export time — the cached value of Qty(A) * Price(B) is written.
    @Test
    public void exportFormula_cellReference_computedAtExport() throws Exception {
        final FormulaRefRow row = new FormulaRefRow();
        row.setQty(6);
        row.setPrice(7);
        row.setTotal("=A2*B2");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("FormulaRef", Arrays.asList(row), FormulaRefRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            final Cell cell = firstDataCell(workbook, "FormulaRef", "Total");
            assertThat(cell.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("A2*B2");
            assertThat(cell.getNumericCellValue()).isEqualTo(42.0);   // 6 * 7
        }
    }

    // Streaming (SXSSF): delegates computation to Excel — only sets the recalc flag and leaves no cached value.
    @Test
    public void exportFormula_sxssf_delegatesRecalcWithoutCaching() throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=2+3");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.SXSSF)
                .exportDataValidation(false)
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Formula", Arrays.asList(row), FormulaRow.class)
                .override(option)
                .toFile(excelFile);

        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            // The flag delegating a full recalculation when the file is opened is set.
            assertThat(workbook.getForceFormulaRecalculation()).isTrue();

            final Cell cell = firstDataCell(workbook, "Formula", "Formula");
            assertThat(cell.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("2+3");
            // Since it is not computed at export time, there is no cached value (0.0). The actual value is filled in by Excel on open.
            assertThat(cell.getNumericCellValue()).isEqualTo(0.0);
        }
    }

    // ------------------------------------------------------------------
    // Null field export: the default exportNullString("") produces an empty-string STRING cell, not a blank cell
    // (type-independent — buildDataCell's null gate applies exportNullString before type dispatch)
    // ------------------------------------------------------------------

    // A null Double is written as a STRING cell containing an empty string, neither numeric 0 nor blank (default exportNullString="").
    @Test
    public void exportNullDouble_writesEmptyStringCell_notBlank() throws Exception {
        final AllTypesRow row = new AllTypesRow();
        row.setText("keep");          // at least one value in the row
        row.setWrapDouble(null);      // target: null Double

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Types", Arrays.asList(row), AllTypesRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            final Cell cell = firstDataCell(workbook, "Types", "WrapDouble");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);   // neither BLANK nor NUMERIC
            assertThat(cell.getStringCellValue()).isEqualTo("");
        }
    }

    // When exportNullString is specified, a null field is written as that string.
    @Test
    public void exportNullString_customValue_overridesNullDoubleCell() throws Exception {
        final AllTypesRow row = new AllTypesRow();
        row.setText("keep");
        row.setWrapDouble(null);

        final PxlExportColumnOption nullOption = PxlExportColumnOption.builder()
                .fieldName("wrapDouble")
                .exportNullString("N/A")
                .build();
        final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                .exportColumnOptions(Arrays.asList(nullOption))
                .build();
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(sheetOption))
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Types", Arrays.asList(row), AllTypesRow.class)
                .override(option)
                .toFile(excelFile);

        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            final Cell cell = firstDataCell(workbook, "Types", "WrapDouble");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(cell.getStringCellValue()).isEqualTo("N/A");
        }
    }

    // ==================================================================
    // Non-XSSF format (HSSF .xls / SXSSF streaming xlsx) feature coverage
    // (extends the type-fidelity/feature verification previously only on the default XSSF to both formats)
    // ==================================================================

    private static PxlExportWorkbookOption formatOption(final PxlFileFormat format, final boolean dataValidation) {
        return PxlExportWorkbookOption.builder()
                .exportFileFormat(format)
                .exportDataValidation(dataValidation)
                .build();
    }

    // Per-format file extension (HSSF=.xls, XSSF/SXSSF=.xlsx). Also prevents file name collisions in parameterized tests.
    private static String ext(final PxlFileFormat format) {
        return "_" + format.name() + (format == PxlFileFormat.HSSF ? ".xls" : ".xlsx");
    }

    // Whether rich-type (all types) round trips are preserved on HSSF and SXSSF too. On SXSSF the auto-size tracking path is also exercised.
    @ParameterizedTest
    @EnumSource(value = PxlFileFormat.class, names = {"HSSF", "SXSSF"})
    public void richTypes_roundTrip_perFormat(final PxlFileFormat format) throws Exception {
        final AllTypesRow row = Fixtures.sampleAllTypesRow();

        final File file = TestPaths.exportFile(testInfo, ext(format));
        pxl.exportExcel()
                .sheet("Types", Arrays.asList(row), AllTypesRow.class)
                .override(formatOption(format, false))
                .toFile(file);

        final AllTypesRow out = pxl.importExcel()
                .sheet(AllTypesRow.class, Arrays.asList("Types"))
                .fromFile(file).get(0);
        Fixtures.assertSampleAllTypesRow(out);
    }

    // Per-format sheet limits match POI SpreadsheetVersion (HSSF=EXCEL97, XSSF/SXSSF=EXCEL2007). Only HSSF differs in value.
    @Test
    public void fileFormat_exportLimits_matchSpreadsheetVersion() {
        assertThat(PxlFileFormat.HSSF.getMaxExportRows()).isEqualTo(65536);
        assertThat(PxlFileFormat.HSSF.getMaxExportColumns()).isEqualTo(256);

        assertThat(PxlFileFormat.XSSF.getMaxExportRows()).isEqualTo(1_048_576);
        assertThat(PxlFileFormat.XSSF.getMaxExportColumns()).isEqualTo(16_384);

        // SXSSF shares the same EXCEL2007 limits as XSSF.
        assertThat(PxlFileFormat.SXSSF.getMaxExportRows()).isEqualTo(PxlFileFormat.XSSF.getMaxExportRows());
        assertThat(PxlFileFormat.SXSSF.getMaxExportColumns()).isEqualTo(PxlFileFormat.XSSF.getMaxExportColumns());
    }

    // A dropdown (exportOptionItems) exports without exception on HSSF/SXSSF too and the value round trips (data validation kept enabled).
    @ParameterizedTest
    @EnumSource(value = PxlFileFormat.class, names = {"HSSF", "SXSSF"})
    public void dropdown_roundTrip_perFormat(final PxlFileFormat format) throws Exception {
        final OptionItemsRow row = new OptionItemsRow();
        row.setChoice("Red");

        final File file = TestPaths.exportFile(testInfo, ext(format));
        pxl.exportExcel()
                .sheet("Opt", Arrays.asList(row), OptionItemsRow.class)
                .override(formatOption(format, true))
                .toFile(file);

        final OptionItemsRow out = pxl.importExcel()
                .sheet(OptionItemsRow.class, Arrays.asList("Opt"))
                .fromFile(file).get(0);
        assertThat(out.getChoice()).isEqualTo("Red");
    }

    // Grouping (splitting sheets by field value) works on HSSF/SXSSF too (the workbook option takes precedence over the class and overrides the format).
    @ParameterizedTest
    @EnumSource(value = PxlFileFormat.class, names = {"HSSF", "SXSSF"})
    public void grouping_splitsIntoSheets_perFormat(final PxlFileFormat format) throws Exception {
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales")));

        final File file = TestPaths.exportFile(testInfo, ext(format));
        pxl.exportExcel()
                .workbook(workbook)
                .override(formatOption(format, false))
                .toFile(file);

        try (Workbook poi = WorkbookFactory.create(file)) {
            assertThat(poi.getSheet("Employees - Engineering")).as("Engineering group sheet").isNotNull();
            assertThat(poi.getSheet("Employees - Sales")).as("Sales group sheet").isNotNull();
        }

        final List<Employee> engineering = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Employees - Engineering"))
                .fromFile(file);
        assertThat(engineering).extracting(Employee::getName).containsExactly("Alice");
    }

    // Sample/template export creates a header row + example row on HSSF (.xls) too.
    @Test
    public void sampleExport_hssf_writesHeaderAndSampleRow() throws Exception {
        final File file = TestPaths.exportFile(testInfo, ".xls");
        pxl.exportSampleExcel()
                .workbook(AllTypesWorkbook.class)
                .override(formatOption(PxlFileFormat.HSSF, false))
                .toFile(file);

        try (Workbook poi = WorkbookFactory.create(file)) {
            final Sheet sheet = poi.getSheet("AllTypes");
            assertThat(sheet).as("AllTypes sheet must be created").isNotNull();
            assertThat(sheet.getRow(0)).as("header row must exist").isNotNull();
            assertThat(sheet.getRow(1)).as("sample row must exist").isNotNull();
        }
    }

    // On SXSSF streaming too, the auto-size path (trackColumnForAutoSizing -> autoSizeColumns) works and column widths fall within the clamp range.
    @Test
    public void sxssf_autoSizeColumn_appliedWithinClamp() throws Exception {
        final File file = TestPaths.exportFile(testInfo, ".xlsx");
        pxl.exportExcel()
                .sheet("Types", Arrays.asList(Fixtures.sampleAllTypesRow()), AllTypesRow.class)
                .override(formatOption(PxlFileFormat.SXSSF, false))
                .toFile(file);

        try (Workbook poi = WorkbookFactory.create(file)) {
            final Cell textCell = firstDataCell(poi, "Types", "Text");
            final int width = poi.getSheet("Types").getColumnWidth(textCell.getColumnIndex());
            // autoSizeColumns clamps to [MIN=2000, MAX=15000] -> being within that range means the auto-size path ran correctly.
            assertThat(width).isBetween(2000, 15000);
        }
    }

    // ------------------------------------------------------------------
    // Grouping export: null group key (a null grouping-field value) and a null row object
    // ------------------------------------------------------------------

    @Test
    public void exportGrouping_nullGroupKeyValue_createsUngroupedSheet() throws Exception {
        // An employee whose grouping field (department) is null lands in the "(ungrouped)" group sheet.
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Dana", 28, "40000", true, LocalDate.of(2021, 5, 1), Grade.C, null)));   // null department

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toFile(excelFile);

        try (Workbook poi = WorkbookFactory.create(excelFile)) {
            assertThat(poi.getSheet("Employees - Engineering")).isNotNull();
            assertThat(poi.getSheet("Employees - (ungrouped)")).as("null grouping value -> (ungrouped) sheet").isNotNull();
        }
    }

    @Test
    public void exportGrouping_nullRowObject_throws() {
        // A null row object cannot be read for its grouping key / cell values and is rejected fail-fast.
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                null));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlCellCodecException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .override(noValidationOption())
                        .toStream(outputStream));
    }
}
