package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.builder.PxlExcelExportBuilder;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Round-trip (export -> import) tests.
 * <p>
 * Exports in workbook / sheet / multi-sheet form, then imports again and verifies that values are preserved.
 * Since the binding/codec logic follows the same path regardless of transport, the transport axis is
 * <b>parameterized</b> via {@link Transport} (real file / in-memory stream / returned POI workbook), running one body
 * across three transports. {@code FILE} mode leaves a {@code <methodName>.xlsx} file on disk, while
 * {@code STREAM}/{@code POI} round-trip in memory.
 */
public class PxlRoundTripTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    // Captures the current test method name to match the (FILE mode) export file name.
    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

    /**
     * Round-trip transport. The codec/binder logic is identical; only the I/O boundary differs.
     */
    public enum Transport {
        /**
         * Export to a real file -> import from the file. (The file remains on disk.)
         */
        FILE,
        /**
         * Export to an in-memory OutputStream -> import from an InputStream.
         */
        STREAM,
        /**
         * Export via the POI Workbook-returning API -> serialize, then import from a stream.
         */
        POI
    }

    /**
     * Holds the artifact (file or bytes) exported by the chosen transport, and reads it back in workbook/sheet form.
     */
    private static final class RoundTripSource {

        private final File file;      // FILE mode
        private final byte[] bytes;   // STREAM / POI mode

        private RoundTripSource(final File file) {
            this.file = file;
            this.bytes = null;
        }

        private RoundTripSource(final byte[] bytes) {
            this.file = null;
            this.bytes = bytes;
        }

        private Object importWorkbook(final String workbookName, final Class<?> workbookClass) throws Exception {
            return file != null
                    ? pxl.importExcel()
                    .workbookName(workbookName)
                    .workbook(workbookClass)
                    .fromFile(file)
                    : pxl.importExcel()
                    .workbookName(workbookName)
                    .workbook(workbookClass)
                    .fromStream(new ByteArrayInputStream(bytes));
        }

        private Collection<?> importSheet(final List<String> candidateSheetNames, final Class<?> rowClass) throws Exception {
            return file != null
                    ? pxl.importExcel()
                    .sheet(rowClass, candidateSheetNames)
                    .fromFile(file)
                    : pxl.importExcel()
                    .sheet(rowClass, candidateSheetNames)
                    .fromStream(new ByteArrayInputStream(bytes));
        }
    }

    // ------------------------------------------------------------------
    // Per-transport export helpers (workbook / sheet / multi-sheet form)
    // ------------------------------------------------------------------

    private RoundTripSource exportWorkbook(final Transport transport,
                                           final Object workbookObject,
                                           final PxlExportWorkbookOption option) throws Exception {
        switch (transport) {
            case FILE: {
                final File file = TestPaths.exportFile(testInfo);
                pxl.exportExcel()
                        .workbook(workbookObject)
                        .override(option)
                        .toFile(file);
                return new RoundTripSource(file);
            }
            case STREAM: {
                final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                pxl.exportExcel()
                        .workbook(workbookObject)
                        .override(option)
                        .toStream(outputStream);
                return new RoundTripSource(outputStream.toByteArray());
            }
            case POI:
            default:
                return new RoundTripSource(toBytes(pxl.exportExcel()
                        .workbook(workbookObject)
                        .override(option)
                        .toWorkbook()));
        }
    }

    private <T> RoundTripSource exportSheet(final Transport transport,
                                            final String sheetName,
                                            final Collection<T> rows,
                                            final Class<T> rowClass,
                                            final PxlExportWorkbookOption option) throws Exception {
        switch (transport) {
            case FILE: {
                final File file = TestPaths.exportFile(testInfo);
                pxl.exportExcel()
                        .sheet(rowClass, rows, sheetName)
                        .override(option)
                        .toFile(file);
                return new RoundTripSource(file);
            }
            case STREAM: {
                final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                pxl.exportExcel()
                        .sheet(rowClass, rows, sheetName)
                        .override(option)
                        .toStream(outputStream);
                return new RoundTripSource(outputStream.toByteArray());
            }
            case POI:
            default:
                return new RoundTripSource(toBytes(pxl.exportExcel()
                        .sheet(rowClass, rows, sheetName)
                        .override(option)
                        .toWorkbook()));
        }
    }

    // Since the builder config methods (sheet/workbook) throw checked Pxl exceptions, this interface is used
    // instead of java.util.function.Consumer (which is non-throwing).
    @FunctionalInterface
    private interface SheetConfigurer {
        void accept(PxlExcelExportBuilder builder) throws PxlException;
    }

    private RoundTripSource exportMultiSheet(final Transport transport,
                                             final SheetConfigurer sheets,
                                             final PxlExportWorkbookOption option) throws Exception {
        switch (transport) {
            case FILE: {
                final File file = TestPaths.exportFile(testInfo);
                final PxlExcelExportBuilder builder = pxl.exportExcel();
                sheets.accept(builder);
                builder.override(option)
                        .toFile(file);
                return new RoundTripSource(file);
            }
            case STREAM: {
                final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                final PxlExcelExportBuilder builder = pxl.exportExcel();
                sheets.accept(builder);
                builder.override(option)
                        .toStream(outputStream);
                return new RoundTripSource(outputStream.toByteArray());
            }
            case POI:
            default: {
                final PxlExcelExportBuilder builder = pxl.exportExcel();
                sheets.accept(builder);
                return new RoundTripSource(toBytes(builder.override(option)
                        .toWorkbook()));
            }
        }
    }

    private static byte[] toBytes(final Workbook workbook) throws Exception {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            return outputStream.toByteArray();
        } finally {
            workbook.close();
        }
    }

    private static List<Employee> sampleEmployees() {
        return Arrays.asList(
                Fixtures.employee("Alice", 30, "50000.50", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000.00", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"));
    }

    private static List<Department> sampleDepartments() {
        return Arrays.asList(
                Fixtures.department("ENG", "Engineering", 12),
                Fixtures.department("SAL", "Sales", 7));
    }

    // ------------------------------------------------------------------
    // Sheet form: all types - transport parity (FILE/STREAM/POI)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(Transport.class)
    public void sheetForm_allTypes_roundTrips(final Transport transport) throws Exception {
        final RoundTripSource source = exportSheet(transport, "AllTypes",
                Arrays.asList(Fixtures.sampleAllTypesRow()), AllTypesRow.class, noValidationOption());

        @SuppressWarnings("unchecked") final List<AllTypesRow> rows =
                (List<AllTypesRow>) source.importSheet(Arrays.asList("AllTypes"), AllTypesRow.class);

        assertThat(rows).hasSize(1);
        Fixtures.assertSampleAllTypesRow(rows.get(0));
    }

    // ------------------------------------------------------------------
    // Workbook form: all types + workbook name - transport parity
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(Transport.class)
    public void workbookForm_allTypes_roundTrips(final Transport transport) throws Exception {
        final AllTypesWorkbook workbook = new AllTypesWorkbook();
        workbook.setWorkbookName("MyWorkbook");
        workbook.setRows(Arrays.asList(Fixtures.sampleAllTypesRow()));

        final RoundTripSource source = exportWorkbook(transport, workbook, noValidationOption());
        final AllTypesWorkbook imported = (AllTypesWorkbook) source.importWorkbook("MyWorkbook", AllTypesWorkbook.class);

        assertThat(PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(imported)).isEqualTo("MyWorkbook");
        assertThat(imported.getRows()).hasSize(1);
        Fixtures.assertSampleAllTypesRow(imported.getRows().get(0));
    }

    // ------------------------------------------------------------------
    // Workbook form: two different sheet types (Employees / Departments) - transport parity
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(Transport.class)
    public void workbookForm_twoSheetTypes_roundTrips(final Transport transport) throws Exception {
        final CompanyWorkbook workbook = new CompanyWorkbook();
        workbook.setWorkbookName("Acme");
        workbook.setEmployees(sampleEmployees());
        workbook.setDepartments(sampleDepartments());

        final RoundTripSource source = exportWorkbook(transport, workbook, noValidationOption());
        final CompanyWorkbook imported = (CompanyWorkbook) source.importWorkbook("Acme", CompanyWorkbook.class);

        assertThat(PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(imported)).isEqualTo("Acme");
        assertThat(imported.getEmployees()).hasSize(2);
        final Employee alice = imported.getEmployees().get(0);
        assertThat(alice.getName()).isEqualTo("Alice");
        assertThat(alice.getAge()).isEqualTo(30);
        assertThat(alice.getSalary()).isEqualByComparingTo("50000.50");
        assertThat(alice.getActive()).isTrue();
        assertThat(alice.getHireDate()).isEqualTo(LocalDate.of(2020, 1, 15));
        assertThat(alice.getGrade()).isEqualTo(Grade.A);
        assertThat(alice.getDepartment()).isEqualTo("Engineering");
        assertThat(imported.getDepartments()).hasSize(2);
        assertThat(imported.getDepartments().get(1).getDepartmentName()).isEqualTo("Sales");
        assertThat(imported.getDepartments().get(1).getHeadcount()).isEqualTo(7);
    }

    // Option null -> exportDataValidation defaults to true (a realistic file with a Grade enum dropdown/hidden sheet); round-trips regardless of transport
    @ParameterizedTest
    @EnumSource(Transport.class)
    public void workbookForm_twoSheetTypesWithValidation_roundTrips(final Transport transport) throws Exception {
        final CompanyWorkbook workbook = new CompanyWorkbook();
        workbook.setWorkbookName("Acme");
        workbook.setEmployees(sampleEmployees());
        workbook.setDepartments(sampleDepartments());

        final RoundTripSource source = exportWorkbook(transport, workbook, null);
        final CompanyWorkbook imported = (CompanyWorkbook) source.importWorkbook("Acme", CompanyWorkbook.class);

        // Even with a dropdown/hidden sheet, name matching reads only the data sheet correctly.
        assertThat(imported.getEmployees()).hasSize(2);
        assertThat(imported.getEmployees().get(0).getGrade()).isEqualTo(Grade.A);
        assertThat(imported.getDepartments()).hasSize(2);
    }

    // ------------------------------------------------------------------
    // Multi-sheet (explicit list) form - transport parity
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(Transport.class)
    public void multiSheetListForm_roundTrips(final Transport transport) throws Exception {
        final List<Employee> engineering = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"));
        final List<Employee> sales = Arrays.asList(
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"),
                Fixtures.employee("Carol", 35, "68000", true, LocalDate.of(2019, 3, 20), Grade.A, "Sales"));

        final RoundTripSource source = exportMultiSheet(transport,
                builder -> builder
                        .sheet(Employee.class, engineering, "Engineering")
                        .sheet(Employee.class, sales, "Sales"),
                noValidationOption());

        @SuppressWarnings("unchecked") final List<Employee> importedEng =
                (List<Employee>) source.importSheet(Arrays.asList("Engineering"), Employee.class);
        @SuppressWarnings("unchecked") final List<Employee> importedSales =
                (List<Employee>) source.importSheet(Arrays.asList("Sales"), Employee.class);

        assertThat(importedEng).extracting(Employee::getName).containsExactly("Alice");
        assertThat(importedSales).extracting(Employee::getName).containsExactly("Bob", "Carol");
    }

    // ------------------------------------------------------------------
    // Real .xls (HSSF) file round-trip (file-format axis - separate from transport parameterization)
    // ------------------------------------------------------------------

    @Test
    public void sheetForm_allTypesViaXlsFile_roundTrips() throws Exception {
        final File file = TestPaths.exportFile(testInfo, ".xls");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.HSSF)
                .exportDataValidation(false)
                .build();
        pxl.exportExcel()
                .sheet(AllTypesRow.class, Arrays.asList(Fixtures.sampleAllTypesRow()), "AllTypes")
                .override(option)
                .toFile(file);

        assertThat(file).exists();

        final List<AllTypesRow> rows = pxl.importExcel()
                .sheet(AllTypesRow.class, Arrays.asList("AllTypes"))
                .fromFile(file);

        assertThat(rows).hasSize(1);
        Fixtures.assertSampleAllTypesRow(rows.get(0));
    }

    // ------------------------------------------------------------------
    // CSV round trip (exportCsv -> importCsv)
    // ------------------------------------------------------------------

    @Test
    public void csvRoundTrip_sheetForm_preservesValues() throws Exception {
        final File csvFile = TestPaths.exportFile(testInfo, ".csv");

        final Employee alice = new Employee();
        alice.setName("Alice");
        alice.setAge(30);
        alice.setSalary(new java.math.BigDecimal("50000.50"));
        alice.setActive(Boolean.TRUE);
        alice.setHireDate(LocalDate.of(2020, 1, 15));
        alice.setGrade(Grade.A);
        alice.setDepartment("Engineering");

        pxl.exportCsv()
                .sheet(Employee.class, Arrays.asList(alice), "Employees")
                .toFile(csvFile);

        final List<Employee> rows = pxl.importCsv()
                .sheet(Employee.class)
                .fromFile(csvFile);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getName()).isEqualTo("Alice");
        assertThat(rows.get(0).getAge()).isEqualTo(30);
        assertThat(rows.get(0).getSalary()).isEqualByComparingTo(new java.math.BigDecimal("50000.50"));
        assertThat(rows.get(0).getActive()).isTrue();
        assertThat(rows.get(0).getHireDate()).isEqualTo(LocalDate.of(2020, 1, 15));
        assertThat(rows.get(0).getGrade()).isEqualTo(Grade.A);
        assertThat(rows.get(0).getDepartment()).isEqualTo("Engineering");
    }

    // ------------------------------------------------------------------
    // Static helpers (PxlWorkbookUtils workbook name / PxlExcelEngine export engine extraction)
    // ------------------------------------------------------------------

    @Test
    public void workbookUtils_getWorkbookName_nullObject_returnsNull() {
        assertThat(PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(null)).isNull();
    }

    @Test
    public void excelEngine_fromWorkbookObject_annotatedAndDefault_resolves() {
        assertThat(PxlExcelEngine.fromWorkbookObject(CompanyWorkbook.class)).isEqualTo(PxlExcelEngine.XSSF);
        assertThat(PxlExcelEngine.fromWorkbookObject(XlsFormatWorkbook.class)).isEqualTo(PxlExcelEngine.HSSF);
    }

    @Test
    public void excelEngine_fromWorkbookObject_nullClass_returnsDefault() {
        assertThat(PxlExcelEngine.fromWorkbookObject(null)).isEqualTo(PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE);
    }
}
