package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.builder.PxlExcelExportBuilder;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlSystemException;
import io.github.hclimkr.pxl.option.*;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.time.*;
import java.util.*;
import java.util.stream.Stream;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static io.github.hclimkr.pxl.tcdata.TestExports.*;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Excel export path tests - masking/trimming/grouping/password/export engine (SXSSF, HSSF), literal formulas, column inheritance, sheet name normalization, lastColumnIndex boundary.
 * <p>
 * Whatever an export produces has to be the same on every terminal the builder offers, so a test whose subject is
 * the exported result is swept across {@link ExportDest} with {@link TestExports#emit} (bytes) or
 * {@link TestExports#workbookOf} (POI workbook) rather than being pinned to one destination. What stays a plain
 * {@code @Test} is the opposite kind: a test whose subject <em>is</em> one destination's mechanics - the caller's
 * stream not being closed, nothing being left on disk when a file cannot be opened, one builder driven through two
 * terminals in a row.
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

    // Reads rows back out of exported bytes, so every destination is verified through the same import call.
    private <T> List<T> importSheet(final byte[] bytes, final Class<T> rowClass, final String sheetName) throws Exception {
        return pxl.importExcel()
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    // As above, for an encrypted export.
    private <T> List<T> importSheet(final byte[] bytes, final Class<T> rowClass, final String sheetName, final String password) throws Exception {
        return pxl.importExcel()
                .override(PxlImportWorkbookOption.builder().importPassword(password).build())
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    // Exports the given rows into a single sheet on the given destination, then imports them back.
    private <T> List<T> roundTripSheet(final ExportDest dest, final String sheetName, final List<T> rows, final Class<T> rowClass) throws Exception {
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(rowClass, rows, sheetName)
                .override(noValidationOption()), dest, testInfo);
        return importSheet(bytes, rowClass, sheetName);
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
    // Builder reuse - running the same builder again with the same configuration
    // ------------------------------------------------------------------

    // Not swept: the subject is two different terminals run off one builder, so the destinations are the fixture.
    @Test
    public void exportExcel_sameBuilderRunTwice_producesIdenticalContent() throws Exception {
        final List<Employee> employees = twoEmployees();

        // The configuration stays on the builder, so each terminal call builds a fresh workbook out of it.
        final PxlExcelExportBuilder builder = pxl.exportExcel()
                .sheet(Employee.class, employees, "People")
                .override(noValidationOption());

        final File excelFile = TestPaths.exportFile(testInfo);
        builder.toFile(excelFile);

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        builder.toStream(outputStream);
        final byte[] secondRun = outputStream.toByteArray();

        // Running a terminal a second time must not add the sheet again.
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(secondRun))) {
            assertThat(workbook.getNumberOfSheets()).isEqualTo(1);
        }
        assertThat(firstSheetName(secondRun)).isEqualTo("People");

        final List<Employee> fromFile = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);
        final List<Employee> fromStream = importSheet(secondRun, Employee.class, "People");

        // Both runs carry exactly the same rows.
        assertThat(fromFile).extracting(Employee::getName).containsExactly("Alice", "Bob");
        assertThat(fromStream).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Terminal failure - the workbook is built before the destination is opened,
    // so a destination that cannot be opened must not leave a file behind or a workbook unreleased
    // ------------------------------------------------------------------

    // A destination whose parent directory does not exist: the failure lands between building the workbook and
    // writing it, which is exactly the window in which the built workbook has to be released anyway.
    private File unopenableFile() {
        return new File(TestPaths.exportFile(testInfo).getPath() + ".no-such-dir", "out.xlsx");
    }

    // Not swept: only the file destination can fail to open.
    @Test
    public void exportExcel_toFileDestinationUnopenable_throwsAndBuilderStaysUsable() throws Exception {
        final PxlExcelExportBuilder builder = pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(noValidationOption());

        final File unopenable = unopenableFile();
        // Opening the file fails with an IOException, which the terminal normalizes.
        assertThrows(PxlSystemException.class, () -> builder.toFile(unopenable));
        // Nothing may be left at the destination - the write never started.
        assertThat(unopenable).doesNotExist();

        // The workbook built for the failed run was released and the configuration is untouched,
        // so the next terminal call builds a fresh workbook out of the same builder.
        final File excelFile = TestPaths.exportFile(testInfo);
        builder.toFile(excelFile);
        assertThat(pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile))
                .extracting(Employee::getName)
                .containsExactly("Alice", "Bob");
    }

    // Same failure on the streaming format, where a workbook left unreleased would also leave POI temp files
    // on disk (closeWorkbook disposes of them).
    @Test
    public void exportSxssf_toFileDestinationUnopenable_throwsAndBuilderStaysUsable() throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.SXSSF)
                .exportDataValidation(false)
                .build();

        final PxlExcelExportBuilder builder = pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(option);

        assertThrows(PxlSystemException.class, () -> builder.toFile(unopenableFile()));

        final File excelFile = TestPaths.exportFile(testInfo);
        builder.toFile(excelFile);
        assertThat(pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile))
                .extracting(Employee::getName)
                .containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Multiple sheets - a different rowClass per sheet
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportMultiSheet_perSheetRowClass_appliesEachClass(final ExportDest dest) throws Exception {
        final List<Employee> employees = twoEmployees();
        final List<AllTypesRow> allTypes = Arrays.asList(Fixtures.sampleAllTypesRow());

        // Call sheet() multiple times to specify a different rowClass per sheet.
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, employees, "Employees")
                .sheet(AllTypesRow.class, allTypes, "AllTypes")
                .override(noValidationOption()), dest, testInfo);

        final List<Employee> importedEmployees = importSheet(bytes, Employee.class, "Employees");
        final List<AllTypesRow> importedAllTypes = importSheet(bytes, AllTypesRow.class, "AllTypes");

        // Verify each sheet is bound with its own rowClass - different classes are applied per sheet.
        assertThat(importedEmployees).extracting(Employee::getName).containsExactly("Alice", "Bob");
        assertThat(importedAllTypes).hasSize(1);
        assertThat(importedAllTypes.get(0).getText()).isEqualTo("Hello, PXL");
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportMultiSheet_threeSheets_preservesCallOrderAndRows(final ExportDest dest) throws Exception {
        final List<Employee> engineering = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"));
        final List<Employee> sales = Arrays.asList(
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"),
                Fixtures.employee("Carol", 35, "68000", true, LocalDate.of(2019, 3, 20), Grade.A, "Sales"));
        final List<Department> departments = Arrays.asList(
                Fixtures.department("ENG", "Engineering", 12),
                Fixtures.department("SAL", "Sales", 7));

        // Three sheet() calls -> three sheets; the same rowClass may repeat and the call order is the sheet order.
        final PxlExcelExportBuilder builder = pxl.exportExcel()
                .sheet(Employee.class, engineering, "Engineering")
                .sheet(Employee.class, sales, "Sales")
                .sheet(Department.class, departments, "Departments")
                .override(noValidationOption());

        final byte[] bytes = emit(builder, dest, testInfo);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            assertThat(workbook.getNumberOfSheets()).isEqualTo(3);
            assertThat(workbook.getSheetName(0)).isEqualTo("Engineering");
            assertThat(workbook.getSheetName(1)).isEqualTo("Sales");
            assertThat(workbook.getSheetName(2)).isEqualTo("Departments");
        }

        // Rows must not leak across sheets: each sheet holds only the rows passed with its own sheet() call.
        final List<Employee> importedEngineering = importSheet(bytes, Employee.class, "Engineering");
        final List<Employee> importedSales = importSheet(bytes, Employee.class, "Sales");
        final List<Department> importedDepartments = importSheet(bytes, Department.class, "Departments");

        assertThat(importedEngineering).extracting(Employee::getName).containsExactly("Alice");
        assertThat(importedSales).extracting(Employee::getName).containsExactly("Bob", "Carol");
        assertThat(importedDepartments).extracting(Department::getCode).containsExactly("ENG", "SAL");
        assertThat(importedDepartments.get(0).getHeadcount()).isEqualTo(12);
    }

    // ------------------------------------------------------------------
    // Masking (exportMasking)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportMasking_digits_replaced(final ExportDest dest) throws Exception {
        final MaskingRow row = new MaskingRow();
        row.setSecret("ID12345");

        final List<MaskingRow> rows = roundTripSheet(dest, "Masking", Arrays.asList(row), MaskingRow.class);

        assertThat(rows).hasSize(1);
        // All 5 digits are replaced with '*'.
        assertThat(rows.get(0).getSecret()).isEqualTo("ID*****");
    }

    // ------------------------------------------------------------------
    // Export trim (exportTrim)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportTrim_whitespace_trimmed(final ExportDest dest) throws Exception {
        final TrimRow row = new TrimRow();
        row.setPadded("  spaced  ");

        final List<TrimRow> rows = roundTripSheet(dest, "Trim", Arrays.asList(row), TrimRow.class);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getPadded()).isEqualTo("spaced");
    }

    // ------------------------------------------------------------------
    // Grouping (exportGroupingFieldName) - sheets are split by department value.
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportGrouping_splitsIntoSheets(final ExportDest dest) throws Exception {
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"),
                Fixtures.employee("Carol", 35, "68000", true, LocalDate.of(2019, 3, 20), Grade.A, "Engineering")));

        final byte[] bytes = emit(pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption()), dest, testInfo);

        // Group sheet names follow the "<sheet name> - <group value>" format.
        try (Workbook poiWorkbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            assertThat(poiWorkbook.getSheet("Employees - Engineering")).as("Engineering group sheet is missing.").isNotNull();
            assertThat(poiWorkbook.getSheet("Employees - Sales")).as("Sales group sheet is missing.").isNotNull();
        }

        final List<Employee> engineering = importSheet(bytes, Employee.class, "Employees - Engineering");
        final List<Employee> sales = importSheet(bytes, Employee.class, "Employees - Sales");

        assertThat(engineering).extracting(Employee::getName).containsExactly("Alice", "Carol");
        assertThat(sales).extracting(Employee::getName).containsExactly("Bob");
    }

    // ------------------------------------------------------------------
    // Password (exportPassword / importPassword) round trip
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportPassword_encrypted_roundTrips(final ExportDest dest) throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();

        // toWorkbook() hands the workbook over unencrypted by contract, so the sweep applies the same password on
        // the way out - which is what the terminal's javadoc tells a caller to do.
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(exportOption), dest, testInfo, XLSX, "secret");

        final List<Employee> people = importSheet(bytes, Employee.class, "People", "secret");

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportPassword_declaredOnAnnotation_encryptsAndReopens(final ExportDest dest) throws Exception {
        // The round trip above drives both passwords from an option; @PxlWorkbook is the other way in.
        final PasswordWorkbook workbook = new PasswordWorkbook();
        workbook.setWorkbookName("Protected");
        workbook.setPeople(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering")));

        final byte[] bytes = emit(pxl.exportExcel()
                .workbook(workbook), dest, testInfo, XLSX, "secret");

        // Opening it without the password has to fail, or the export would have written plaintext.
        assertThrows(PxlException.class, () -> importSheet(bytes, Employee.class, "People"));

        final PasswordWorkbook imported = pxl.importExcel()
                .workbookName("Protected")
                .workbook(PasswordWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(imported.getPeople()).extracting(Employee::getName).containsExactly("Alice");
    }

    // ------------------------------------------------------------------
    // Engine: SXSSF (streaming write, result is xlsx)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportSxssf_streaming_roundTrips(final ExportDest dest) throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.SXSSF)
                .exportDataValidation(false)
                .build();

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(option), dest, testInfo);

        final List<Employee> people = importSheet(bytes, Employee.class, "People");

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
    }

    // ------------------------------------------------------------------
    // Engine: HSSF (xls) round trip
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportHssf_xls_roundTrips(final ExportDest dest) throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.HSSF)
                .exportDataValidation(false)
                .build();

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(option), dest, testInfo, XLS);

        final List<Employee> people = importSheet(bytes, Employee.class, "People");

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
        assertThat(people.get(0).getGrade()).isEqualTo(Grade.A);
    }

    // ------------------------------------------------------------------
    // Engine x password: SXSSF/HSSF encrypted round trip (writeToStream uses agile encryption for both XSSF and SXSSF, Biff8 for HSSF)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportSxssf_encrypted_roundTrips(final ExportDest dest) throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.SXSSF)
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(exportOption), dest, testInfo, XLSX, "secret");

        final List<Employee> people = importSheet(bytes, Employee.class, "People", "secret");

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");
        assertThat(people.get(0).getGrade()).isEqualTo(Grade.A);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportHssf_encrypted_roundTrips(final ExportDest dest) throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");

        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.HSSF)
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(exportOption), dest, testInfo, XLS, "secret");

        final List<Employee> people = importSheet(bytes, Employee.class, "People", "secret");

        assertThat(people).hasSize(1);
        assertThat(people.get(0).getName()).isEqualTo("Alice");

        // The Biff8 thread-local key is cleaned up in finally, so a subsequent unencrypted HSSF export is not contaminated.
        final PxlExportWorkbookOption plainOption = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.HSSF)
                .exportDataValidation(false)
                .build();
        final byte[] plainBytes = emit(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(plainOption), dest, testInfo, "-plain" + XLS);

        assertThat(importSheet(plainBytes, Employee.class, "People")).hasSize(1);
    }

    // ------------------------------------------------------------------
    // Resource ownership: toStream does not close the OutputStream passed by the caller (only flushes)
    // ------------------------------------------------------------------

    // Not swept: only the stream destination is handed a stream it does not own.
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
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(noValidationOption())
                .toStream(tracking);

        assertThat(closed[0]).as("toStream must not close the caller's stream").isFalse();
        assertThat(tracking.size()).isGreaterThan(0);
    }

    // ------------------------------------------------------------------
    // exportStringAsFormula=false: "=..." strings are not evaluated as formulas but preserved as literals
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportLiteralFormula_notEvaluated(final ExportDest dest) throws Exception {
        final LiteralRow row = new LiteralRow();
        row.setExpr("=1+2");

        final List<LiteralRow> out = roundTripSheet(dest, "Literal", Arrays.asList(row), LiteralRow.class);

        assertThat(out).hasSize(1);
        // If evaluated as a formula it would become "3", but as a literal it must stay "=1+2".
        assertThat(out.get(0).getExpr()).isEqualTo("=1+2");
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportLiteralFormula_writtenQuotePrefixed(final ExportDest dest) throws Exception {
        // The round-trip above shows the value survives; this pins how. Without exportStringAsFormula the leading '='
        // makes the cell a quote-prefixed string, so Excel shows the text verbatim instead of reading it as a formula.
        // That prefix is a safeguard on writing text, not one of the forms the column options choose between.
        final LiteralRow row = new LiteralRow();
        row.setExpr("=1+2");

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(LiteralRow.class, Arrays.asList(row), "Literal")
                .override(noValidationOption()), dest, testInfo)) {
            final Cell cell = firstDataCell(workbook, "Literal", "Expr");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(cell.getStringCellValue()).isEqualTo("=1+2");
            assertThat(cell.getCellStyle().getQuotePrefixed()).isTrue();
        }
    }

    // ------------------------------------------------------------------
    // Column inheritance: @PxlColumn fields from the superclass are also bound
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportInheritedColumns_roundTrips(final ExportDest dest) throws Exception {
        final DerivedRow row = new DerivedRow();
        row.setId(7);                 // inherited field
        row.setBaseName("base");      // inherited field
        row.setExtra("own");          // own field

        final List<DerivedRow> out = roundTripSheet(dest, "Derived", Arrays.asList(row), DerivedRow.class);

        assertThat(out).hasSize(1);
        assertThat(out.get(0).getId()).isEqualTo(7);
        assertThat(out.get(0).getBaseName()).isEqualTo("base");
        assertThat(out.get(0).getExtra()).isEqualTo("own");
    }

    // Separately verify that a superclass-only field actually appears in the header
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportInheritedColumns_headerIncludesThem(final ExportDest dest) throws Exception {
        final DerivedRow row = new DerivedRow();
        row.setId(1);
        row.setBaseName("base");
        row.setExtra("own");

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(DerivedRow.class, Arrays.asList(row), "Derived")
                .override(noValidationOption()), dest, testInfo)) {
            final Row header = workbook.getSheet("Derived").getRow(0);
            final Set<String> headers = new HashSet<>();
            for (final Cell cell : header) {
                headers.add(cell.getStringCellValue());
            }
            assertThat(headers).contains("Id", "BaseName", "Extra");
        }
    }

    // ------------------------------------------------------------------
    // All sheets exportEnabled=false: fails because there is no sheet to export
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportAllSheetsDisabled_throws(final ExportDest dest) {
        final DisabledSheetWorkbook workbook = new DisabledSheetWorkbook();
        workbook.setWorkbookName("NoSheets");
        workbook.setRows(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering")));

        // The check sits in build(), which every terminal runs first, so the failure is the same on all of them.
        assertThrows(PxlDataException.class, () ->
                emit(pxl.exportExcel()
                        .workbook(workbook)
                        .override(noValidationOption()), dest, testInfo));
    }

    // ------------------------------------------------------------------
    // Empty password -> round trips without encryption (importable without a password)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportEmptyPassword_noEncryption_roundTrips(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportPassword("")     // empty password = no encryption
                .build();

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(exportOption), dest, testInfo, XLSX, "");

        // Must be importable without any password option
        assertThat(importSheet(bytes, Employee.class, "People"))
                .extracting(Employee::getName)
                .containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Import with password + stream reader combination
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportPassword_withStreamReader_roundTrips(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(exportOption), dest, testInfo, XLSX, "secret");

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
                .fromStream(new ByteArrayInputStream(bytes));
        assertThat(people).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Sheet name over 31 chars -> truncated to 31 chars
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportSheetName_over31Chars_truncated(final ExportDest dest) throws Exception {
        final String longName = "VeryLongSheetNameThatExceeds31Characters";   // 40 chars
        assertThat(longName.length()).isGreaterThan(31);

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), longName)
                .override(noValidationOption()), dest, testInfo);

        // The actual sheet name is truncated to 31 chars
        final String actualName = firstSheetName(bytes);
        assertThat(actualName).hasSize(31);
        assertThat(longName).startsWith(actualName);

        // Importing with the truncated actual name round trips correctly
        assertThat(importSheet(bytes, Employee.class, actualName))
                .extracting(Employee::getName)
                .containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Invalid characters in sheet name -> replaced with spaces
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportSheetName_invalidChars_sanitized(final ExportDest dest) throws Exception {
        // Invalid characters in Excel sheet names: / \ ? * [ ] :
        final String badName = "Bad/Name:With*Chars";

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), badName)
                .override(noValidationOption()), dest, testInfo);

        final String actualName = firstSheetName(bytes);
        // No invalid characters must remain
        assertThat(actualName).doesNotContain("/", "\\", "?", "*", "[", "]", ":");
        // Invalid characters are replaced with spaces
        assertThat(actualName).isEqualTo("Bad Name With Chars");

        // Importing with the normalized name round trips correctly
        assertThat(importSheet(bytes, Employee.class, actualName))
                .extracting(Employee::getName)
                .containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Exception when the export column range is smaller than the column count
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportLastColumnIndex_tooSmall_throws(final ExportDest dest) {
        final ExportColBoundWorkbook workbook = new ExportColBoundWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering")));

        assertThrows(PxlDataException.class, () ->
                emit(pxl.exportExcel()
                        .workbook(workbook)
                        .override(PxlExportWorkbookOption.builder().exportDataValidation(false).build()), dest, testInfo));
    }

    // ------------------------------------------------------------------
    // Date/time numeric export (when neither pattern nor masking is set, written as numeric date cells rather than strings)
    // ------------------------------------------------------------------

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

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportExcel_dateTimeColumnsWithoutPatternOrMasking_writtenAsNumericDateCells(final ExportDest dest) throws Exception {
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

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(DateTimeNumericRow.class, Arrays.asList(row), "DateTimes")
                .override(noValidationOption()), dest, testInfo);

        final String[] dateHeaders = {"JavaDate", "LocalDate", "LocalTime", "LocalDateTime", "ZonedDateTime", "OffsetTime", "OffsetDateTime"};
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
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
        final List<DateTimeNumericRow> imported = importSheet(bytes, DateTimeNumericRow.class, "DateTimes");
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

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportExcel_dateColumnWithMasking_staysStringCell(final ExportDest dest) throws Exception {
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

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(option), dest, testInfo)) {
            final Cell cell = firstDataCell(workbook, "People", "HireDate");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(cell.getStringCellValue()).isEqualTo("****-**-**");   // all digits of 2020-01-15 are masked
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportExcel_dateColumnWithExportPattern_staysStringCell(final ExportDest dest) throws Exception {
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

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(option), dest, testInfo)) {
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
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportFormula_nonStreaming_cachesComputedResult(final ExportDest dest) throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=2+3");

        // Open with raw POI but do not evaluate here - getNumericCellValue() returns the cached value left by export.
        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(FormulaRow.class, Arrays.asList(row), "Formula")
                .override(noValidationOption()), dest, testInfo)) {
            final Cell cell = firstDataCell(workbook, "Formula", "Formula");
            assertThat(cell.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("2+3");
            // If export had not computed it, the cached value would be 0.0 - 5.0 means it was computed at export time.
            assertThat(cell.getNumericCellValue()).isEqualTo(5.0);
        }
    }

    // exportStringAsFormula only claims a value that starts with '=': anything else is ordinary text.
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportFormula_valueWithoutLeadingEquals_writesPlainText(final ExportDest dest) throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("not a formula");

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(FormulaRow.class, Arrays.asList(row), "Formula")
                .override(noValidationOption()), dest, testInfo)) {
            final Cell cell = firstDataCell(workbook, "Formula", "Formula");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(cell.getStringCellValue()).isEqualTo("not a formula");
            // no leading '=', so there is nothing for the quote prefix to guard against either
            assertThat(cell.getCellStyle().getQuotePrefixed()).isFalse();
        }
    }

    // A cell-reference formula (=A2*B2) is also computed at export time - the cached value of Qty(A) * Price(B) is written.
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportFormula_cellReference_computedAtExport(final ExportDest dest) throws Exception {
        final FormulaRefRow row = new FormulaRefRow();
        row.setQty(6);
        row.setPrice(7);
        row.setTotal("=A2*B2");

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(FormulaRefRow.class, Arrays.asList(row), "FormulaRef")
                .override(noValidationOption()), dest, testInfo)) {
            final Cell cell = firstDataCell(workbook, "FormulaRef", "Total");
            assertThat(cell.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("A2*B2");
            assertThat(cell.getNumericCellValue()).isEqualTo(42.0);   // 6 * 7
        }
    }

    // Streaming (SXSSF): delegates computation to Excel - only sets the recalc flag and leaves no cached value.
    // Swept over the written destinations only: on WORKBOOK the assertion would run against the live SXSSF
    // workbook, whose formula cells carry no readable value at all rather than the absent one written to a file.
    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportFormula_sxssf_delegatesRecalcWithoutCaching(final ExportDest dest) throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=2+3");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.SXSSF)
                .exportDataValidation(false)
                .build();

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(FormulaRow.class, Arrays.asList(row), "Formula")
                .override(option), dest, testInfo)) {
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
    // (type-independent - buildDataCell's null gate applies exportNullString before type dispatch)
    // ------------------------------------------------------------------

    // A null Double is written as a STRING cell containing an empty string, neither numeric 0 nor blank (default exportNullString="").
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportNullDouble_writesEmptyStringCell_notBlank(final ExportDest dest) throws Exception {
        final AllTypesRow row = new AllTypesRow();
        row.setText("keep");          // at least one value in the row
        row.setWrapDouble(null);      // target: null Double

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(AllTypesRow.class, Arrays.asList(row), "Types")
                .override(noValidationOption()), dest, testInfo)) {
            final Cell cell = firstDataCell(workbook, "Types", "WrapDouble");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);   // neither BLANK nor NUMERIC
            assertThat(cell.getStringCellValue()).isEqualTo("");
        }
    }

    // When exportNullString is specified, a null field is written as that string.
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportNullString_customValue_overridesNullDoubleCell(final ExportDest dest) throws Exception {
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

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(AllTypesRow.class, Arrays.asList(row), "Types")
                .override(option), dest, testInfo)) {
            final Cell cell = firstDataCell(workbook, "Types", "WrapDouble");
            assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(cell.getStringCellValue()).isEqualTo("N/A");
        }
    }

    // ==================================================================
    // Non-XSSF engine (HSSF .xls / SXSSF streaming xlsx) feature coverage
    // (extends the type-fidelity/feature verification previously only on the default XSSF to both engines)
    // ==================================================================

    private static PxlExportWorkbookOption engineOption(final PxlExcelEngine engine, final boolean dataValidation) {
        return PxlExportWorkbookOption.builder()
                .exportExcelEngine(engine)
                .exportDataValidation(dataValidation)
                .build();
    }

    // Per-engine file extension (HSSF=.xls, XSSF/SXSSF=.xlsx), taken from the format the engine writes.
    // Also prevents file name collisions between the engines of one parameterized test (the destination adds its own).
    private static String ext(final PxlExcelEngine engine) {
        return "_" + engine.name() + "." + engine.getFileFormat().getFilenameExtension();
    }

    // The engine x destination matrix the sweeps below run: two axes, so an @EnumSource on either alone would
    // leave the other pinned to a single value.
    private static Stream<Arguments> nonDefaultEnginesAndDestinations() {
        final List<Arguments> arguments = new ArrayList<>();
        for (final PxlExcelEngine engine : Arrays.asList(PxlExcelEngine.HSSF, PxlExcelEngine.SXSSF)) {
            for (final ExportDest dest : ExportDest.values()) {
                arguments.add(Arguments.of(engine, dest));
            }
        }
        return arguments.stream();
    }

    // Whether rich-type (all types) round trips are preserved on HSSF and SXSSF too. On SXSSF the auto-size tracking path is also exercised.
    @ParameterizedTest
    @MethodSource("nonDefaultEnginesAndDestinations")
    public void richTypes_roundTrip_perEngine(final PxlExcelEngine engine, final ExportDest dest) throws Exception {
        final AllTypesRow row = Fixtures.sampleAllTypesRow();

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(AllTypesRow.class, Arrays.asList(row), "Types")
                .override(engineOption(engine, false)), dest, testInfo, ext(engine));

        Fixtures.assertSampleAllTypesRow(importSheet(bytes, AllTypesRow.class, "Types").get(0));
    }

    // Per-format sheet limits match POI SpreadsheetVersion (XLS=EXCEL97, XLSX=EXCEL2007). Only XLS differs in value.
    @Test
    public void fileFormat_exportLimits_matchSpreadsheetVersion() {
        assertThat(PxlFileFormat.XLS.getMaxExportRows()).isEqualTo(65536);
        assertThat(PxlFileFormat.XLS.getMaxExportColumns()).isEqualTo(256);

        assertThat(PxlFileFormat.XLSX.getMaxExportRows()).isEqualTo(1_048_576);
        assertThat(PxlFileFormat.XLSX.getMaxExportColumns()).isEqualTo(16_384);
    }

    // Each engine writes exactly one physical format, and the limits it is bound by are that format's.
    // XSSF and SXSSF differ in memory behaviour only, so they share both the format and the limits.
    @Test
    public void excelEngine_fileFormat_mapsToWrittenFormatAndItsLimits() {
        assertThat(PxlExcelEngine.HSSF.getFileFormat()).isEqualTo(PxlFileFormat.XLS);
        assertThat(PxlExcelEngine.XSSF.getFileFormat()).isEqualTo(PxlFileFormat.XLSX);
        assertThat(PxlExcelEngine.SXSSF.getFileFormat()).isEqualTo(PxlFileFormat.XLSX);

        assertThat(PxlExcelEngine.SXSSF.getFileFormat().getMaxExportRows())
                .isEqualTo(PxlExcelEngine.XSSF.getFileFormat().getMaxExportRows());
        assertThat(PxlExcelEngine.SXSSF.getFileFormat().getMaxExportColumns())
                .isEqualTo(PxlExcelEngine.XSSF.getFileFormat().getMaxExportColumns());
    }

    // A dropdown (exportOptionItems) exports without exception on HSSF/SXSSF too and the value round trips (data validation kept enabled).
    @ParameterizedTest
    @MethodSource("nonDefaultEnginesAndDestinations")
    public void dropdown_roundTrip_perEngine(final PxlExcelEngine engine, final ExportDest dest) throws Exception {
        final OptionItemsRow row = new OptionItemsRow();
        row.setChoice("Red");

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(OptionItemsRow.class, Arrays.asList(row), "Opt")
                .override(engineOption(engine, true)), dest, testInfo, ext(engine));

        assertThat(importSheet(bytes, OptionItemsRow.class, "Opt").get(0).getChoice()).isEqualTo("Red");
    }

    // Grouping (splitting sheets by field value) works on HSSF/SXSSF too (the workbook option takes precedence over the class and overrides the engine).
    @ParameterizedTest
    @MethodSource("nonDefaultEnginesAndDestinations")
    public void grouping_splitsIntoSheets_perEngine(final PxlExcelEngine engine, final ExportDest dest) throws Exception {
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales")));

        final byte[] bytes = emit(pxl.exportExcel()
                .workbook(workbook)
                .override(engineOption(engine, false)), dest, testInfo, ext(engine));

        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            assertThat(poi.getSheet("Employees - Engineering")).as("Engineering group sheet").isNotNull();
            assertThat(poi.getSheet("Employees - Sales")).as("Sales group sheet").isNotNull();
        }

        assertThat(importSheet(bytes, Employee.class, "Employees - Engineering"))
                .extracting(Employee::getName)
                .containsExactly("Alice");
    }

    // Sample/template export creates a header row + example row on HSSF (.xls) too.
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sampleExport_hssf_writesHeaderAndSampleRow(final ExportDest dest) throws Exception {
        try (Workbook poi = workbookOf(pxl.exportSampleExcel()
                .workbook(AllTypesWorkbook.class)
                .override(engineOption(PxlExcelEngine.HSSF, false)), dest, testInfo, XLS)) {
            final Sheet sheet = poi.getSheet("AllTypes");
            assertThat(sheet).as("AllTypes sheet must be created").isNotNull();
            assertThat(sheet.getRow(0)).as("header row must exist").isNotNull();
            assertThat(sheet.getRow(1)).as("sample row must exist").isNotNull();
        }
    }

    // On SXSSF streaming too, the auto-size path (trackColumnForAutoSizing -> autoSizeColumns) works and column widths fall within the clamp range.
    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sxssf_autoSizeColumn_appliedWithinClamp(final ExportDest dest) throws Exception {
        try (Workbook poi = workbookOf(pxl.exportExcel()
                .sheet(AllTypesRow.class, Arrays.asList(Fixtures.sampleAllTypesRow()), "Types")
                .override(engineOption(PxlExcelEngine.SXSSF, false)), dest, testInfo)) {
            final Cell textCell = firstDataCell(poi, "Types", "Text");
            final int width = poi.getSheet("Types").getColumnWidth(textCell.getColumnIndex());
            // autoSizeColumns clamps to [MIN=2000, MAX=15000] -> being within that range means the auto-size path ran correctly.
            assertThat(width).isBetween(2000, 15000);
        }
    }

    // ------------------------------------------------------------------
    // Grouping export: null group key (a null grouping-field value) and a null row object
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportGrouping_nullGroupKeyValue_createsUngroupedSheet(final ExportDest dest) throws Exception {
        // An employee whose grouping field (department) is null lands in the "(ungrouped)" group sheet.
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Dana", 28, "40000", true, LocalDate.of(2021, 5, 1), Grade.C, null)));   // null department

        try (Workbook poi = workbookOf(pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption()), dest, testInfo)) {
            assertThat(poi.getSheet("Employees - Engineering")).isNotNull();
            assertThat(poi.getSheet("Employees - (ungrouped)")).as("null grouping value -> (ungrouped) sheet").isNotNull();
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportGrouping_nullRowObject_throwsNamingPosition(final ExportDest dest) {
        // The grouping branch reads a field off every row to build its key, so it meets a null row before the
        // write loop does. It is rejected by the same up-front check, with the same message.
        final GroupedWorkbook workbook = new GroupedWorkbook();
        workbook.setWorkbookName("Grouped");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                null));

        final PxlDataException exception = assertThrows(PxlDataException.class, () ->
                emit(pxl.exportExcel()
                        .workbook(workbook)
                        .override(noValidationOption()), dest, testInfo));

        assertThat(exception.getMessage()).contains("Employees").contains("2");
    }

    // ------------------------------------------------------------------
    // A null row object names its position rather than an innocent column
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportExcel_nullRowObject_throwsNamingPosition(final ExportDest dest) {
        // Reading a field off a null row raises a NullPointerException, which used to arrive as a
        // PxlCellCodecException tagged with the first column - a column that did nothing wrong, since no codec ran.
        // The failure now says what is actually wrong and where, and no cell-codec error is reported.
        final List<Employee> employees = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                null,
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"));

        final PxlDataException exception = assertThrows(PxlDataException.class, () ->
                emit(pxl.exportExcel()
                        .sheet(Employee.class, employees, "Employees")
                        .override(noValidationOption()), dest, testInfo));

        assertThat(exception.getMessage())
                .as("names the sheet and the one-based position of the null element")
                .contains("Employees")
                .contains("2");
        assertThat(exception).isNotInstanceOf(PxlCellCodecException.class);
        assertThat(exception.getMessage())
                .as("no column should be blamed")
                .doesNotContain("Name");
    }

    // Not swept: the guarantee under test is about the file system, which only the file destination touches.
    @Test
    public void exportExcel_nullRowObject_writesNothing() {
        // The check runs before the destination is opened, so the failure leaves no partial file behind -
        // the same guarantee the builder's prepare()/writeTo() ordering gives every other export failure.
        final List<Employee> employees = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                null);

        final File excelFile = TestPaths.exportFile(testInfo);
        assertThrows(PxlDataException.class, () ->
                pxl.exportExcel()
                        .sheet(Employee.class, employees, "Employees")
                        .override(noValidationOption())
                        .toFile(excelFile));

        assertThat(excelFile).as("a failed export must not leave a file behind").doesNotExist();
    }
}
