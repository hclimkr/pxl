package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.util.*;

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
                    .contains("Text", "PrimInt", "Grade", "Category", "Point", "Money", "StringList", "Uuid");
            // The example row is populated with exportSample values.
            assertThat(sampleValuesOf(workbook, "AllTypes"))
                    .containsEntry("Text", "Sample text")
                    .containsEntry("LeadingZero", "007")
                    .containsEntry("Grade", "A")
                    .containsEntry("Category", "Electronics")
                    .containsEntry("Point", "3,7")
                    .containsEntry("Money", "USD 1050")
                    .containsEntry("StringList", "Apple;Banana;Cherry")
                    .containsEntry("Uuid", "123e4567-e89b-12d3-a456-426614174000");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSampleToWorkbook_fromSheetClass_writesHeaders() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(Employee.class, "Sample")
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
                .sheet(Employee.class, "Sample")
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
                .sheet(Employee.class, "Sample")
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
                .sheet(SampleColumnRow.class, "S")
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
    // Multi-sheet sample - calling sheet() multiple times (a different class per sheet)
    // ------------------------------------------------------------------

    @Test
    public void exportSampleMultiSheet_perSheetRowClass_writesEachSheet() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(Employee.class, "People")
                .sheet(SampleColumnRow.class, "Cols")
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
                .sheet(Employee.class, "People")
                .sheet(Department.class, "Depts")
                .sheet(SampleColumnRow.class, "Cols")
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

    // ------------------------------------------------------------------
    // exportSample goes through the same codec as a data value, so the export forms apply to the sample row
    // ------------------------------------------------------------------

    @Test
    public void exportSample_formulaColumn_writesFormulaCell() throws Exception {
        // A sample value is a raw string handed to the String codec, so exportStringAsFormula turns the sample row's
        // cell into a formula rather than into the text "=1+2".
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(FormulaPictureSampleRow.class, "S")
                .toWorkbook();
        try {
            final Cell cell = workbook.getSheet("S").getRow(1).getCell(0);
            assertThat(cell.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("1+2");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_pictureColumn_takesPicturePath() throws Exception {
        // The same applies to exportStringAsPicture: the sample value is treated as an image location. "photo.png" is
        // not one that can be loaded, so it is skipped the way any bad location is and the cell is left empty - the
        // sample row does not fall back to showing the location as text.
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(FormulaPictureSampleRow.class, "S")
                .toWorkbook();
        try {
            final Cell cell = workbook.getSheet("S").getRow(1).getCell(1);
            assertThat(cell.getCellType()).isEqualTo(CellType.BLANK);
            assertThat(workbook.getAllPictures()).isEmpty();
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_javaDateSampleWithTrailingGarbage_throws() {
        // A Date column's exportSample is a raw string that the export formatter has to parse, so it is the export-side
        // counterpart of the import defect: "2020-01-15 xxx" used to be read up to the space and written as
        // 2020-01-15. The pattern must now match the sample in full (issue M2 fix).
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlCellCodecException.class, () ->
                pxl.exportSampleExcel()
                        .sheet(BadJavaDateSampleRow.class, "S")
                        .toStream(outputStream));
    }

    @Test
    public void exportSample_enumSampleNotAnEnumConstant_throws() {
        // An enum column's exportSample is internally reverse-parsed to the enum and re-exported, so it must be a valid value.
        // "N/A", which is not in Grade(A/B/C/F) -> the sample export fails with an exception.
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlCellCodecException.class, () ->
                pxl.exportSampleExcel()
                        .sheet(BadEnumSampleRow.class, "S")
                        .toStream(outputStream));
    }

    @Test
    public void exportSample_uuidSample_writesCanonicalValue() throws Exception {
        // A UUID column's exportSample is parsed back into a UUID and written out again, so it survives unchanged.
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(UuidRow.class, "S")
                .toWorkbook();
        try {
            assertThat(sampleValuesOf(workbook, "S"))
                    .containsEntry("Id", "123e4567-e89b-12d3-a456-426614174000");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_uuidSampleNotCanonical_throws() {
        // "1-1-1-1-1" is a sample UUID.fromString would accept, widening it into another value. The codec refuses it,
        // so the sample export fails instead of writing a value its own import side would reject.
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlCellCodecException.class, () ->
                pxl.exportSampleExcel()
                        .sheet(BadUuidSampleRow.class, "S")
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // Structure: a sample sheet has exactly a header row (0) + one sample row (1), nothing more
    // ------------------------------------------------------------------

    @Test
    public void exportSample_generatesExactlyHeaderRowAndOneSampleRow() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(Employee.class, "Sample")
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
                .sheet(Employee.class, "Sample")
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
                .sheet(SampleColumnRow.class, "S")
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
                .sheet(SampleMixedRow.class, "S")
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
    public void exportSample_workbookClassWithHssfEngine_producesXlsWorkbook() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(XlsFormatWorkbook.class)
                .toWorkbook();
        try {
            assertThat(workbook).as("class-level exportExcelEngine=HSSF -> XLS workbook").isInstanceOf(HSSFWorkbook.class);
            assertThat(headerValuesOf(workbook, "Employees")).contains(EMPLOYEE_HEADERS);
            assertEmployeeSampleValues(workbook, "Employees");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_optionOverridesEngineToHssf_producesXlsWorkbook() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(Employee.class, "Sample")
                .override(PxlExportWorkbookOption.builder().exportExcelEngine(PxlExcelEngine.HSSF).build())
                .toWorkbook();
        try {
            assertThat(workbook).as("option exportExcelEngine=HSSF applies to the sample").isInstanceOf(HSSFWorkbook.class);
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
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportSampleExcel()
                        .workbook(AllTypesWorkbook.class)
                        .sheet(Employee.class, "S")
                        .toStream(outputStream));
    }

    @Test
    public void exportSample_neitherWorkbookNorSheetSpecified_throws() {
        // Neither the workbook form nor the sheet form was configured before the terminal.
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleExcel()
                .toStream(outputStream));
    }

    @Test
    public void exportSample_blankSheetName_throws() {
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleExcel()
                .sheet(Employee.class, "  "));
    }

    @Test
    public void exportSample_nullRowClass_throws() {
        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleExcel()
                .sheet(null, "S"));
    }

    @Test
    public void exportSample_nullWorkbookClass_throws() {
        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleExcel()
                .workbook(null));
    }

    // ------------------------------------------------------------------
    // The sample row is always written, so exportLastDataRowIndex must not shrink the ranges around it
    // ------------------------------------------------------------------

    @Test
    public void exportSample_lastDataRowIndexBeforeFirstDataRow_keepsFilterAndDropdownOverSampleRow() throws Exception {
        // A sample sheet carries exactly one data row whatever the declared bound says. With the header on 0-based
        // row 0 the sample lands on row 1, so a declared bound of 1 (1-based) points at the header row -- ahead of
        // the row actually written. The auto-filter and the dropdown have to follow the written row regardless.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .exportColumnFilter(true)
                        .exportLastDataRowIndex(1)
                        .build()))
                .build();

        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(SampleDropdownRow.class, "Sample")
                .override(option)
                .toWorkbook();
        try {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Sample");

            // Two columns, header on row 1 and the sample on row 2 (1-based, as the reference writes it).
            assertThat(sheet.getCTWorksheet().getAutoFilter().getRef())
                    .as("the auto filter should span the header and the sample row")
                    .isEqualTo("A1:B2");
            assertThat(sheet.getDataValidations())
                    .as("the dropdown should still be attached to the sample row")
                    .hasSize(1);
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSample_lastDataRowIndexBeforeHeaderRow_doesNotThrow() throws Exception {
        // The header is pushed to 1-based row 3, so the sample lands on row 4 while the declared bound points at
        // row 1. A range built from that bound would end before it starts, which POI rejects outright.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .exportColumnFilter(true)
                        .exportHeaderRowIndex(3)
                        .exportLastDataRowIndex(1)
                        .build()))
                .build();

        final Workbook workbook = pxl.exportSampleExcel()
                .sheet(SampleDropdownRow.class, "Sample")
                .override(option)
                .toWorkbook();
        try {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Sample");

            assertThat(sheet.getCTWorksheet().getAutoFilter().getRef())
                    .as("the auto filter should span the header and the sample row")
                    .isEqualTo("A3:B4");

            // The workbook has to survive being written out, which is where an invalid range would surface.
            final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            assertThat(outputStream.size()).isPositive();
        } finally {
            workbook.close();
        }
    }
}
