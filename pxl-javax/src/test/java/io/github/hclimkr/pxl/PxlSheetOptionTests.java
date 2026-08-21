package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.option.*;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static io.github.hclimkr.pxl.tcdata.TestExports.emit;
import static io.github.hclimkr.pxl.tcdata.TestExports.workbookOf;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Tests for various @PxlSheet property combinations.
 * <p>
 * Verifies header/data row and column indices (export+import) and the guards on them, hidden row/column exclusion,
 * exportColumnFilter, sheet-level exportEnabled/exportSampleEnabled/importEnabled, sheet names (aliases name={...}
 * and the fall back to the field name), importLastDataRowIndex bounds, and exportRowHeightInPoints.
 * <p>
 * Every test that exports is swept across {@link ExportDest}: what a sheet option writes has to be the same on
 * every terminal. The import-only and option-accessor tests never reach a terminal at all.
 */
public class PxlSheetOptionTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

    private byte[] exportWorkbookBytes(final Object workbook, final ExportDest dest) throws Exception {
        return emit(pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption()), dest, testInfo);
    }

    private static List<Employee> twoEmployees() {
        return Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"));
    }

    // ------------------------------------------------------------------
    // Header/data row and column indices (export + import) round-trip
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetIndexShift_customIndices_roundTrips(final ExportDest dest) throws Exception {
        final IndexShiftWorkbook workbook = new IndexShiftWorkbook();
        workbook.setWorkbookName("Shifted");
        workbook.setData(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(workbook, dest);

        // verify export positions: header row is 1-based 3 = 0-based 2, first data column is 1-based 2 = 0-based 1
        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            final Sheet sheet = poi.getSheet("Data");
            assertThat(sheet.getRow(0)).as("empty row (0) above the header row").isNull();
            final Row header = sheet.getRow(2);
            assertThat(header).as("header row should be at 0-based 2").isNotNull();
            assertThat(header.getCell(0)).as("first column (0) should be empty").isNull();
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("Name");
        }

        // import reads with the same indices to round-trip
        final IndexShiftWorkbook imported = pxl.importExcel()
                .workbookName("Shifted")
                .workbook(IndexShiftWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(imported.getData()).hasSize(2);
        assertThat(imported.getData()).extracting(Employee::getName).containsExactly("Alice", "Bob");
        assertThat(imported.getData().get(0).getDepartment()).isEqualTo("Engineering");
    }

    // ------------------------------------------------------------------
    // Hidden row exclusion (importExcludeHiddenRows)
    // ------------------------------------------------------------------

    private static byte[] buildSheetWithHiddenRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("Data");
            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Department");

            final String[][] data = {{"Alice", "Engineering"}, {"Bob", "Sales"}, {"Carol", "Finance"}};
            for (int r = 0; r < data.length; r++) {
                final Row row = sheet.createRow(r + 1);
                row.createCell(0).setCellValue(data[r][0]);
                row.createCell(1).setCellValue(data[r][1]);
            }
            // hide the 2nd data row (Bob)
            sheet.getRow(2).setZeroHeight(true);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static List<Employee> importData(final byte[] bytes, final PxlImportWorkbookOption option) throws Exception {
        return pxl.importExcel()
                .override(option)
                .sheet(Employee.class, Arrays.asList("Data"))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    @Test
    public void excludeHiddenRows_hiddenRowsSkipped() throws Exception {
        final byte[] bytes = buildSheetWithHiddenRow();

        // default (false): hidden rows included -> 3 people
        assertThat(importData(bytes, null)).extracting(Employee::getName).containsExactly("Alice", "Bob", "Carol");

        // exclude=true: hidden row (Bob) excluded -> 2 people
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder().importExcludeHiddenRows(true).build()))
                .build();
        assertThat(importData(bytes, option)).extracting(Employee::getName).containsExactly("Alice", "Carol");
    }

    // ------------------------------------------------------------------
    // Hidden column exclusion (importExcludeHiddenColumns)
    // ------------------------------------------------------------------

    private static byte[] buildSheetWithHiddenColumn() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("Data");
            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Department");

            final Row row = sheet.createRow(1);
            row.createCell(0).setCellValue("Alice");
            row.createCell(1).setCellValue("Engineering");

            // hide Department (column 1)
            sheet.setColumnHidden(1, true);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    @Test
    public void excludeHiddenColumns_hiddenColumnsSkipped() throws Exception {
        final byte[] bytes = buildSheetWithHiddenColumn();

        // default: Department included
        assertThat(importData(bytes, null).get(0).getDepartment()).isEqualTo("Engineering");

        // exclude=true: hidden column (Department) excluded -> null
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder().importExcludeHiddenColumns(true).build()))
                .build();
        final Employee alice = importData(bytes, option).get(0);
        assertThat(alice.getName()).isEqualTo("Alice");
        assertThat(alice.getDepartment()).isNull();
    }

    private static byte[] buildSheetWithHiddenRowAndColumn() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("Data");
            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Department");

            final String[][] data = {{"Alice", "Engineering"}, {"Bob", "Sales"}, {"Carol", "Finance"}};
            for (int r = 0; r < data.length; r++) {
                final Row row = sheet.createRow(r + 1);
                row.createCell(0).setCellValue(data[r][0]);
                row.createCell(1).setCellValue(data[r][1]);
            }
            // hide the 2nd data row (Bob) and the Department column
            sheet.getRow(2).setZeroHeight(true);
            sheet.setColumnHidden(1, true);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    @Test
    public void excludeHidden_declaredOnAnnotation_skipsHiddenRowAndColumn() throws Exception {
        // The two tests above drive these flags from a runtime option; the annotation is the other way in, and
        // nothing else covers it.
        final byte[] bytes = buildSheetWithHiddenRowAndColumn();

        final HiddenSheetWorkbook imported = pxl.importExcel()
                .workbookName("Hidden")
                .workbook(HiddenSheetWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(imported.getRows()).extracting(Employee::getName).containsExactly("Alice", "Carol");
        assertThat(imported.getRows()).extracting(Employee::getDepartment).containsOnlyNulls();
    }

    // ------------------------------------------------------------------
    // exportColumnFilter (auto filter)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportColumnFilter_autoFilterApplied(final ExportDest dest) throws Exception {
        final ColumnFilterWorkbook workbook = new ColumnFilterWorkbook();
        workbook.setWorkbookName("Filter");
        workbook.setRows(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(workbook, dest);
        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            final XSSFSheet sheet = (XSSFSheet) poi.getSheet("Filtered");
            assertThat(sheet.getCTWorksheet().isSetAutoFilter()).as("auto filter should be set").isTrue();
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportColumnFilter_notRequested_noAutoFilter(final ExportDest dest) throws Exception {
        // Without this the test above would pass just as well if every sheet were filtered regardless of the flag.
        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(noValidationOption()), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("People");
            assertThat(sheet.getCTWorksheet().isSetAutoFilter()).as("the flag defaults to false").isFalse();
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportColumnFilter_sheetOptionFalse_overridesAnnotation(final ExportDest dest) throws Exception {
        final ColumnFilterWorkbook workbook = new ColumnFilterWorkbook();
        workbook.setWorkbookName("Filter");
        workbook.setRows(twoEmployees());

        // The sheet declares exportColumnFilter=true; the option is read first and takes it back.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .exportColumnFilter(false)
                        .build()))
                .build();

        try (Workbook poi = workbookOf(pxl.exportExcel()
                .workbook(workbook)
                .override(option), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) poi.getSheet("Filtered");
            assertThat(sheet.getCTWorksheet().isSetAutoFilter()).as("the option outranks the annotation").isFalse();
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportColumnFilter_sheetFormOption_spansHeaderAndDataRows(final ExportDest dest) throws Exception {
        // The sheet form carries no @PxlSheet, so the option is the only source of the flag. Employee binds seven
        // columns and two rows follow the header, which puts the filter over A1:G3.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .exportColumnFilter(true)
                        .build()))
                .build();

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(option), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("People");
            assertThat(sheet.getCTWorksheet().getAutoFilter().getRef())
                    .as("the filter should span the header row and both data rows")
                    .isEqualTo("A1:G3");
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportColumnFilter_shiftedIndices_followsHeaderRowAndFirstColumn(final ExportDest dest) throws Exception {
        // With the header pushed to 1-based row 3 and the columns to 1-based column 2, the filter has to start
        // where the sheet actually begins (B3) rather than at A1, and end on the last written row and column (H5).
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .exportColumnFilter(true)
                        .exportHeaderRowIndex(3)
                        .exportFirstDataColumnIndex(2)
                        .build()))
                .build();

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(option), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("People");
            assertThat(sheet.getCTWorksheet().getAutoFilter().getRef()).isEqualTo("B3:H5");
        }
    }

    // ------------------------------------------------------------------
    // Sheet-level exportSampleEnabled
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetSampleDisabled_excludedFromSample(final ExportDest dest) throws Exception {
        try (Workbook workbook = workbookOf(pxl.exportSampleExcel()
                .workbook(SampleToggleWorkbook.class), dest, testInfo)) {
            assertThat(workbook.getSheet("WithSample")).as("sheet included in sample").isNotNull();
            assertThat(workbook.getSheet("NoSample")).as("exportSampleEnabled=false sheet is excluded from sample").isNull();
        }
    }

    // ------------------------------------------------------------------
    // Sheet-level importEnabled
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetImportDisabled_notImported(final ExportDest dest) throws Exception {
        final ImportToggleWorkbook source = new ImportToggleWorkbook();
        source.setWorkbookName("Toggle");
        source.setEnabled(twoEmployees());
        source.setDisabled(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(source, dest);

        final ImportToggleWorkbook imported = pxl.importExcel()
                .workbookName("Toggle")
                .workbook(ImportToggleWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(imported.getEnabled()).as("enabled sheet is imported").hasSize(2);
        assertThat(imported.getDisabled()).as("importEnabled=false sheet is not imported").isNull();
    }

    // ------------------------------------------------------------------
    // importLastDataRowIndex bounds
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void importLastDataRowIndex_limitsRows(final ExportDest dest) throws Exception {
        // export 3 people
        final List<Employee> three = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, null, Grade.B, "Sales"),
                Fixtures.employee("Carol", 35, "68000", true, null, Grade.A, "Finance"));
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, three, "People")
                .override(noValidationOption()), dest, testInfo);

        // 1-based: header row 1, data 2/3/4. importLastDataRowIndex=3 -> only data 2..3 (Alice, Bob)
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importLastDataRowIndex(3)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<Employee> rows = pxl.importExcel()
                .override(option)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // exportRowHeightInPoints
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportRowHeight_customHeight_applied(final ExportDest dest) throws Exception {
        final RowHeightWorkbook workbook = new RowHeightWorkbook();
        workbook.setWorkbookName("Heights");
        workbook.setRows(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(workbook, dest);
        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            final Sheet sheet = poi.getSheet("Tall");
            // the data row (0-based 1) height should be near 40pt
            assertThat(sheet.getRow(1).getHeightInPoints()).isCloseTo(40.0f, org.assertj.core.data.Offset.offset(0.6f));
        }
    }

    // ------------------------------------------------------------------
    // Sheet option: column-option list accessors + sheet-name normalization
    // ------------------------------------------------------------------

    @Test
    public void importSheetOption_columnOptionAccessors() throws PxlNullPointerException {
        final PxlImportColumnOption c0 = PxlImportColumnOption.builder().fieldName("a").build();
        final PxlImportColumnOption c1 = PxlImportColumnOption.builder().fieldName("b").build();
        final PxlImportSheetOption option = PxlImportSheetOption.builder().build();

        assertThat(option.addImportColumnOption(c0)).isTrue();
        option.addImportColumnOption(c1);

        assertThat(option.getImportColumnOption(0)).isSameAs(c0);
        assertThat(option.getImportColumnOption(9)).isNull();

        assertThat(PxlImportSheetOption.getImportColumnOption(option, 1)).isSameAs(c1);
        assertThat(PxlImportSheetOption.getImportColumnOption(null, 0)).isNull();
        assertThat(PxlImportSheetOption.getImportColumnOptions(option)).containsExactly(c0, c1);
        assertThat(PxlImportSheetOption.getImportColumnOptions(null)).isEmpty();

        assertThrows(PxlNullPointerException.class, () -> option.addImportColumnOption(null));
    }

    @Test
    public void exportSheetOption_columnOptionAccessors() throws PxlNullPointerException {
        final PxlExportColumnOption c0 = PxlExportColumnOption.builder().fieldName("a").build();
        final PxlExportColumnOption c1 = PxlExportColumnOption.builder().fieldName("b").build();
        final PxlExportSheetOption option = PxlExportSheetOption.builder().build();

        assertThat(option.addExportColumnOption(c0)).isTrue();
        option.addExportColumnOption(c1);

        assertThat(option.getExportColumnOption(0)).isSameAs(c0);
        assertThat(option.getExportColumnOption(9)).isNull();

        assertThat(PxlExportSheetOption.getExportColumnOption(option, 1)).isSameAs(c1);
        assertThat(PxlExportSheetOption.getExportColumnOption(null, 0)).isNull();
        assertThat(PxlExportSheetOption.getExportColumnOptions(option)).containsExactly(c0, c1);
        assertThat(PxlExportSheetOption.getExportColumnOptions(null)).isEmpty();

        assertThrows(PxlNullPointerException.class, () -> option.addExportColumnOption(null));
    }

    @Test
    public void importSheetOption_getImportSheetNames_stripsAllWhitespaceAndBlanks() {
        // Import normalization deletes ALL whitespace (inner too) and drops blank entries.
        final PxlImportSheetOption option = PxlImportSheetOption.builder()
                .importSheetNames(Arrays.asList(" A B ", "", "  ", "c")).build();
        assertThat(PxlImportSheetOption.getImportSheetNames(option)).containsExactly("AB", "c");
        assertThat(PxlImportSheetOption.getImportSheetNames(null)).isNull();
        assertThat(PxlImportSheetOption.getImportSheetNames(PxlImportSheetOption.builder().build())).isNull();   // names == null -> null
    }

    @Test
    public void exportSheetOption_getExportSheetNames_trimsAndDropsBlanks() {
        // Export normalization only trims the ends (inner whitespace kept) and drops blank entries.
        final PxlExportSheetOption option = PxlExportSheetOption.builder()
                .exportSheetNames(Arrays.asList(" A B ", "", "c ")).build();
        assertThat(PxlExportSheetOption.getExportSheetNames(option)).containsExactly("A B", "c");
        assertThat(PxlExportSheetOption.getExportSheetNames(null)).isNull();
        assertThat(PxlExportSheetOption.getExportSheetNames(PxlExportSheetOption.builder().build())).isNull();
    }

    // ------------------------------------------------------------------
    // Explicit exportLastDataRowIndex caps the number of written data rows
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportLastDataRowIndex_capsWrittenRows(final ExportDest dest) throws Exception {
        // A non-default (explicit) exportLastDataRowIndex exercises the 1-based -> 0-based bound calculation
        // and limits how many data rows are actually written, below the number of row objects.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder().exportLastDataRowIndex(2).build()))
                .build();
        final List<Employee> employees = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales"),
                Fixtures.employee("Carol", 35, "68000", true, LocalDate.of(2019, 3, 20), Grade.A, "Finance"));

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(Employee.class, employees, "People")
                .override(option), dest, testInfo);

        // header at 0-based 0, first data row at 0-based 1, exportLastDataRowIndex=2 -> exclusive bound 2 -> only row 1
        final List<Employee> imported = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromStream(new ByteArrayInputStream(bytes));
        assertThat(imported).extracting(Employee::getName).containsExactly("Alice");
    }

    // ------------------------------------------------------------------
    // name(): the field name stands in when none is declared; only the first alias is written
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetName_notDeclared_usesFieldName(final ExportDest dest) throws Exception {
        final FieldNameSheetWorkbook workbook = new FieldNameSheetWorkbook();
        workbook.setWorkbookName("ByFieldName");
        workbook.setEmployees(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(workbook, dest);

        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            assertThat(poi.getSheet("employees")).as("the field name becomes the sheet name").isNotNull();
        }

        // The same fallback has to drive the matching on the way back in, or the round trip breaks.
        final FieldNameSheetWorkbook imported = pxl.importExcel()
                .workbookName("ByFieldName")
                .workbook(FieldNameSheetWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(imported.getEmployees()).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetName_aliasArray_exportsUnderTheFirstName(final ExportDest dest) throws Exception {
        // Import accepts any of the aliases, but a sheet can only be written under one of them: the first.
        final AliasSheetWorkbook workbook = new AliasSheetWorkbook();
        workbook.setWorkbookName("Aliased");
        workbook.setData(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(workbook, dest);
        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            assertThat(poi.getSheetName(0)).isEqualTo("Crew");
            assertThat(poi.getSheet("Employee")).as("the remaining aliases are not written as sheets").isNull();
        }
    }

    // ------------------------------------------------------------------
    // Sheet row/column index guards (export). The import side of the same guards lives in PxlExcelImportTests.
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportSheetIndices_invalidCombination_throws(final ExportDest dest) {
        // A negative index.
        assertThrows(PxlDataException.class, () -> exportWithSheetOption(dest, PxlExportSheetOption.builder()
                .exportHeaderRowIndex(-1)
                .build()));

        // The first data row must come after the header row.
        assertThrows(PxlDataException.class, () -> exportWithSheetOption(dest, PxlExportSheetOption.builder()
                .exportHeaderRowIndex(3)
                .exportFirstDataRowIndex(2)
                .build()));

        // The row range must not be inverted.
        assertThrows(PxlDataException.class, () -> exportWithSheetOption(dest, PxlExportSheetOption.builder()
                .exportFirstDataRowIndex(5)
                .exportLastDataRowIndex(3)
                .build()));

        // Neither must the column range.
        assertThrows(PxlDataException.class, () -> exportWithSheetOption(dest, PxlExportSheetOption.builder()
                .exportFirstDataColumnIndex(5)
                .exportLastDataColumnIndex(3)
                .build()));
    }

    private void exportWithSheetOption(final ExportDest dest, final PxlExportSheetOption sheetOption) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(sheetOption))
                .build();

        emit(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(option), dest, testInfo);
    }

    // ------------------------------------------------------------------
    // Sheet-level exportEnabled / importEnabled through a runtime option (the annotation paths are covered above)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetExportDisabled_sheetOption_excludesSheet(final ExportDest dest) throws Exception {
        final SampleToggleWorkbook workbook = new SampleToggleWorkbook();
        workbook.setWorkbookName("Toggle");
        workbook.setWithSample(twoEmployees());
        workbook.setNoSample(twoEmployees());

        // Keyed on the field name, so it reaches one sheet and leaves the other alone.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .fieldName("noSample")
                        .exportEnabled(false)
                        .build()))
                .build();

        try (Workbook poiWorkbook = workbookOf(pxl.exportExcel()
                .workbook(workbook)
                .override(option), dest, testInfo)) {
            assertThat(poiWorkbook.getSheet("WithSample")).isNotNull();
            assertThat(poiWorkbook.getSheet("NoSample")).as("exportEnabled=false from the option").isNull();
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetImportDisabled_sheetOptionEnables_isImported(final ExportDest dest) throws Exception {
        final ImportToggleWorkbook source = new ImportToggleWorkbook();
        source.setWorkbookName("Toggle");
        source.setEnabled(twoEmployees());
        source.setDisabled(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(source, dest);

        // The sheet declares importEnabled=false; the option puts it back.
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .fieldName("disabled")
                        .importEnabled(true)
                        .build()))
                .build();

        final ImportToggleWorkbook imported = pxl.importExcel()
                .workbookName("Toggle")
                .override(option)
                .workbook(ImportToggleWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(imported.getEnabled()).hasSize(2);
        assertThat(imported.getDisabled()).as("the option outranks importEnabled=false").hasSize(2);
    }

    // ------------------------------------------------------------------
    // exportSampleEnabled governs the sample export alone
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void sheetSampleDisabled_dataExport_stillWritesSheet(final ExportDest dest) throws Exception {
        // The counterpart of sheetSampleDisabled_excludedFromSample: the flag has no say over a data export.
        final SampleToggleWorkbook workbook = new SampleToggleWorkbook();
        workbook.setWorkbookName("Toggle");
        workbook.setWithSample(twoEmployees());
        workbook.setNoSample(twoEmployees());

        final byte[] bytes = exportWorkbookBytes(workbook, dest);
        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            assertThat(poi.getSheet("WithSample")).isNotNull();
            assertThat(poi.getSheet("NoSample")).as("exportSampleEnabled=false is not exportEnabled=false").isNotNull();
        }
    }
}
