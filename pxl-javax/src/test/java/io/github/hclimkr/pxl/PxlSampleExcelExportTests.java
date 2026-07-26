package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Sample export tests.
 * <p>
 * Covers the {@link Pxl} sample-family methods, which generate a header row plus one example (sample) row
 * filled from each column's exportSample value. Reopens the generated workbook with POI and verifies that
 * the sheet name, header-row column names, and example-row values are correct.
 */
public class PxlSampleExcelExportTests {

    private static Pxl pxl;

    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

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

    // Returns the set of cell strings in the given sheet's header row (0-based index 0).
    private static Set<String> headerValuesOf(final Workbook workbook, final String sheetName) {
        final Sheet sheet = workbook.getSheet(sheetName);
        assertThat(sheet).as("sheet '" + sheetName + "' was not created").isNotNull();

        final Row headerRow = sheet.getRow(0);
        assertThat(headerRow).as("header row was not created").isNotNull();

        final Set<String> headers = new HashSet<>();
        for (final Cell cell : headerRow) {
            headers.add(cell.getStringCellValue());
        }
        return headers;
    }

    // Pairs the header row (0) and example row (1) by column index and returns a {header -> rendered example value} map.
    private static Map<String, String> sampleValuesOf(final Workbook workbook, final String sheetName) {
        final Sheet sheet = workbook.getSheet(sheetName);
        assertThat(sheet).as("sheet '" + sheetName + "' was not created").isNotNull();

        final Row headerRow = sheet.getRow(0);
        final Row sampleRow = sheet.getRow(1);
        assertThat(sampleRow).as("sample row was not created").isNotNull();

        final Map<String, String> values = new HashMap<>();
        for (final Cell headerCell : headerRow) {
            final int col = headerCell.getColumnIndex();
            final Cell valueCell = sampleRow.getCell(col);
            final String rendered = (valueCell == null) ? "" : DATA_FORMATTER.formatCellValue(valueCell);
            values.put(headerCell.getStringCellValue(), rendered);
        }
        return values;
    }

    private static final String[] EMPLOYEE_HEADERS =
            {"Name", "Age", "Salary", "Active", "HireDate", "Grade", "Department"};

    // Verifies the example-row values of the Employee sheet sample (shared).
    private static void assertEmployeeSampleValues(final Workbook workbook, final String sheetName) {
        assertThat(sampleValuesOf(workbook, sheetName))
                .containsEntry("Name", "Alice")
                .containsEntry("Age", "30")
                .containsEntry("Grade", "A")
                .containsEntry("Department", "Engineering")
                .containsEntry("Active", "true");   // Boolean true -> rendered via default exportTrueString ("true")
    }

    // ------------------------------------------------------------------
    // Sample returning a POI workbook
    // ------------------------------------------------------------------

    @Test
    public void exportSampleToWorkbook_fromWorkbookClass_writesHeaders() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(AllTypesWorkbook.class)
                .toWorkbook();
        try {
            assertThat(headerValuesOf(workbook, "AllTypes"))
                    .contains("Text", "PrimInt", "Grade", "Category", "Point", "Money", "StringList");
            // The example row is populated with exportSample values.
            assertThat(sampleValuesOf(workbook, "AllTypes"))
                    .containsEntry("Text", "Sample text")
                    .containsEntry("LeadingZero", "007")
                    .containsEntry("Grade", "A")
                    .containsEntry("Category", "Electronics")
                    .containsEntry("Point", "3,7")
                    .containsEntry("Money", "USD 1050")
                    .containsEntry("StringList", "Apple;Banana;Cherry");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSampleToWorkbook_fromSheetClass_writesHeaders() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("Sample", Employee.class)
                .toWorkbook();
        try {
            assertThat(headerValuesOf(workbook, "Sample")).contains(EMPLOYEE_HEADERS);
            assertEmployeeSampleValues(workbook, "Sample");
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // File sample
    // ------------------------------------------------------------------

    @Test
    public void exportSampleToFile_fromWorkbookClass_writesSampleValues() throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportSampleExcel()
                .workbook(AllTypesWorkbook.class)
                .toFile(excelFile);

        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            assertThat(headerValuesOf(workbook, "AllTypes")).contains(
                    "Text", "BigInt", "Duration", "Period",
                    "JavaDate", "LocalDate", "LocalTime", "LocalDateTime",
                    "ZonedDateTime", "OffsetTime", "OffsetDateTime");
            assertThat(sampleValuesOf(workbook, "AllTypes"))
                    .containsEntry("Text", "Sample text")
                    .containsEntry("BigInt", "12345678901234567890")
                    .containsEntry("Duration", "PT1H2M3S")
                    .containsEntry("Period", "P1Y2M3D")
                    .containsEntry("IntList", "10;20;30")
                    // Patterned date/time columns are written as string cells whose value equals the exportSample
                    // rendered by that pattern (deterministic, locale-independent).
                    .containsEntry("JavaDate", "2023-06-15 10:30:45")
                    .containsEntry("LocalDate", "2023-06-15")
                    .containsEntry("LocalTime", "10:30:45")
                    .containsEntry("LocalDateTime", "2023-06-15 10:30:45")
                    // Pattern-less ZonedDateTime/OffsetTime/OffsetDateTime are written as numeric date cells (local
                    // part only) with fixed POI builtin formats, so DataFormatter renders them the same in any locale:
                    // datetime = "m/d/yy h:mm" -> "6/15/23 10:30", time = "h:mm:ss" -> "10:30:45".
                    .containsEntry("ZonedDateTime", "6/15/23 10:30")
                    .containsEntry("OffsetTime", "10:30:45")
                    .containsEntry("OffsetDateTime", "6/15/23 10:30");
        }
    }

    @Test
    public void exportSampleToFile_fromSheetClass_writesSampleValues() throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportSampleExcel()
                .sheet("Sample", Employee.class)
                .toFile(excelFile);

        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            assertThat(headerValuesOf(workbook, "Sample")).contains(EMPLOYEE_HEADERS);
            assertEmployeeSampleValues(workbook, "Sample");
        }
    }

    // ------------------------------------------------------------------
    // Stream sample
    // ------------------------------------------------------------------

    @Test
    public void exportSampleToStream_fromWorkbookClass_writesHeaders() throws Exception {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportSampleExcel()
                .workbook(AllTypesWorkbook.class)
                .toStream(outputStream);

        try (InputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            assertThat(headerValuesOf(workbook, "AllTypes")).contains("Text", "Grade");
            assertThat(sampleValuesOf(workbook, "AllTypes"))
                    .containsEntry("Text", "Sample text")
                    .containsEntry("Grade", "A");
        }
    }

    @Test
    public void exportSampleToStream_fromSheetClass_writesHeaders() throws Exception {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportSampleExcel()
                .sheet("Sample", Employee.class)
                .toStream(outputStream);

        try (InputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            assertThat(headerValuesOf(workbook, "Sample")).contains(EMPLOYEE_HEADERS);
            assertEmployeeSampleValues(workbook, "Sample");
        }
    }

    // ------------------------------------------------------------------
    // Column exportSampleEnabled = false
    // ------------------------------------------------------------------

    @Test
    public void columnExportSampleDisabled_excludedFromSample() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("S", SampleColumnRow.class)
                .toWorkbook();
        try {
            final Set<String> headerSet = headerValuesOf(workbook, "S");
            assertThat(headerSet).contains("Keep");
            assertThat(headerSet).doesNotContain("Skip");   // excluded via exportSampleEnabled=false
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Multi-sheet sample — calling sheet() multiple times (a different class per sheet)
    // ------------------------------------------------------------------

    @Test
    public void exportSampleMultiSheet_perSheetRowClass_writesEachSheet() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("People", Employee.class)
                .sheet("Cols", SampleColumnRow.class)
                .toWorkbook();
        try {
            assertThat(workbook.getNumberOfSheets()).isEqualTo(2);

            // First sheet: Employee headers + example row
            assertThat(headerValuesOf(workbook, "People")).contains(EMPLOYEE_HEADERS);
            assertEmployeeSampleValues(workbook, "People");

            // Second sheet: SampleColumnRow (a different class applied per sheet)
            final Set<String> cols = headerValuesOf(workbook, "Cols");
            assertThat(cols).contains("Keep");
            assertThat(cols).doesNotContain("Skip");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSampleMultiSheet_threeSheets_preservesCallOrder() throws Exception {
        // Three sheet() calls -> three sheets, created in call order.
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("People", Employee.class)
                .sheet("Depts", Department.class)
                .sheet("Cols", SampleColumnRow.class)
                .toWorkbook();
        try {
            assertThat(workbook.getNumberOfSheets()).isEqualTo(3);
            assertThat(workbook.getSheetName(0)).isEqualTo("People");
            assertThat(workbook.getSheetName(1)).isEqualTo("Depts");
            assertThat(workbook.getSheetName(2)).isEqualTo("Cols");

            // Each sheet's header row comes from its own row class.
            assertThat(headerValuesOf(workbook, "People")).contains(EMPLOYEE_HEADERS);
            assertThat(headerValuesOf(workbook, "Depts")).contains("Code", "DepartmentName", "Headcount");
            assertThat(headerValuesOf(workbook, "Cols")).contains("Keep");
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // The exportSample of an enum column must be a value parseable to that enum (constraint)
    // ------------------------------------------------------------------

    @Test
    public void exportSample_enumSampleNotAnEnumConstant_throws() {
        // An enum column's exportSample is internally reverse-parsed to the enum and re-exported, so it must be a valid value.
        // "N/A", which is not in Grade(A/B/C/F) -> the sample export fails with an exception.
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlException.class, () ->
                pxl.exportSampleExcel()
                        .sheet("S", BadEnumSampleRow.class)
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // Structure: a sample sheet has exactly a header row (0) + one sample row (1), nothing more
    // ------------------------------------------------------------------

    @Test
    public void exportSample_generatesExactlyHeaderRowAndOneSampleRow() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("Sample", Employee.class)
                .toWorkbook();
        try {
            final Sheet sheet = workbook.getSheet("Sample");
            assertThat(sheet.getRow(0)).as("header row at index 0").isNotNull();
            assertThat(sheet.getRow(1)).as("single sample row at index 1").isNotNull();
            assertThat(sheet.getLastRowNum()).as("no rows beyond the sample row").isEqualTo(1);
            assertThat(sheet.getRow(2)).as("there is no second data row").isNull();
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_everyColumn_getsASampleValue() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("Sample", Employee.class)
                .toWorkbook();
        try {
            final Map<String, String> values = sampleValuesOf(workbook, "Sample");
            assertEmployeeSampleValues(workbook, "Sample");
            // Decimal/date columns also get a sample value (exact rendering is format-dependent, so only non-blank is asserted).
            assertThat(values.get("Salary")).as("Salary sample value").isNotBlank();
            assertThat(values.get("HireDate")).as("HireDate sample value").isNotBlank();
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Sheet-level exportSampleEnabled = false -> the sheet is excluded from the sample
    // ------------------------------------------------------------------

    @Test
    public void exportSample_sheetSampleDisabled_excludesThatSheet() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(SampleToggleWorkbook.class)
                .toWorkbook();
        try {
            assertThat(workbook.getSheet("WithSample")).as("enabled sheet is included").isNotNull();
            assertThat(workbook.getSheet("NoSample")).as("exportSampleEnabled=false sheet is excluded").isNull();
            assertThat(workbook.getNumberOfSheets()).isEqualTo(1);
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Column sample values: present column writes its exportSample; a column without exportSample renders blank
    // ------------------------------------------------------------------

    @Test
    public void exportSample_enabledColumn_writesItsSampleValue() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("S", SampleColumnRow.class)
                .toWorkbook();
        try {
            assertThat(sampleValuesOf(workbook, "S")).containsEntry("Keep", "K");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_columnWithoutExportSample_rendersBlankCell() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("S", SampleMixedRow.class)
                .toWorkbook();
        try {
            assertThat(headerValuesOf(workbook, "S")).contains("Filled", "Empty");
            assertThat(sampleValuesOf(workbook, "S"))
                    .containsEntry("Filled", "V")
                    .containsEntry("Empty", "");   // no exportSample -> blank sample cell
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // File format HSSF (XLS): from the class annotation and from an option override
    // ------------------------------------------------------------------

    @Test
    public void exportSample_workbookClassWithHssfFormat_producesXlsWorkbook() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(XlsFormatWorkbook.class)
                .toWorkbook();
        try {
            assertThat(workbook).as("class-level exportFileFormat=HSSF -> XLS workbook").isInstanceOf(HSSFWorkbook.class);
            assertThat(headerValuesOf(workbook, "Employees")).contains(EMPLOYEE_HEADERS);
            assertEmployeeSampleValues(workbook, "Employees");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_optionOverridesFormatToHssf_producesXlsWorkbook() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet("Sample", Employee.class)
                .override(PxlExportWorkbookOption.builder().exportFileFormat(PxlFileFormat.HSSF).build())
                .toWorkbook();
        try {
            assertThat(workbook).as("option exportFileFormat=HSSF applies to the sample").isInstanceOf(HSSFWorkbook.class);
            assertThat(headerValuesOf(workbook, "Sample")).contains(EMPLOYEE_HEADERS);
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Builder guards
    // ------------------------------------------------------------------

    @Test
    public void exportSample_bothWorkbookAndSheetSpecified_throws() {
        // workbook(Class) and sheet(...) are mutually exclusive; the terminal build fails.
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlException.class, () ->
                pxl.exportSampleExcel()
                        .workbook(AllTypesWorkbook.class)
                        .sheet("S", Employee.class)
                        .toStream(outputStream));
    }

    @Test
    public void exportSample_neitherWorkbookNorSheetSpecified_throws() {
        // Neither the workbook form nor the sheet form was configured before the terminal.
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlException.class, () -> pxl.exportSampleExcel()
                .toStream(outputStream));
    }

    @Test
    public void exportSample_blankSheetName_throws() {
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleExcel()
                .sheet("  ", Employee.class));
    }

    @Test
    public void exportSample_nullRowClass_throws() {
        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleExcel()
                .sheet("S", null));
    }

    @Test
    public void exportSample_nullWorkbookClass_throws() {
        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleExcel()
                .workbook(null));
    }
}
