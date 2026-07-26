package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.builder.PxlExcelImportBuilder;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Set;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Excel/XLS import path tests — NUMERIC cells, merged regions, inherited sheets, sheet name mapping, external XLS, row index boundaries, streaming/formula import.
 */
public class PxlExcelImportTests {

    private static Pxl pxl;

    private static final ZoneId ZONE = ZoneId.systemDefault();
    private static final LocalDateTime BASE = LocalDateTime.of(2023, 6, 15, 10, 30, 45);
    private static final Date BASE_DATE = Date.from(BASE.atZone(ZONE).toInstant());

    private static final String[] HEADERS = {
            "LocalDate", "LocalTime", "LocalDateTime", "JavaDate", "ZonedDateTime", "OffsetDateTime",
            "IntVal", "LongVal", "DoubleVal", "BigDec", "BigInt", "NumericAsString", "CharFromNumeric",
            "BoolFromBoolean", "BoolFromNumeric", "DurationFromNumeric", "PeriodFromNumeric"
    };

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

    private static Date dateOf(final LocalDate localDate) {
        return Date.from(localDate.atStartOfDay(ZONE).toInstant());
    }

    // Creates a NUMERIC date cell with a date format applied. (DateUtil.isCellDateFormatted == true)
    private static void setDateCell(final Row row, final int col, final Date date, final CellStyle dateStyle) {
        final Cell cell = row.createCell(col);
        cell.setCellValue(date);
        cell.setCellStyle(dateStyle);
    }

    /**
     * Builds a workbook composed of NUMERIC/BOOLEAN cells and returns it as bytes.
     * <p>
     * Row 0: header / Row 1: positive/true set / Row 2: negative/false set
     */
    private static byte[] buildNumericFixture() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            final Sheet sheet = workbook.createSheet("Numeric");

            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));

            // header
            final Row header = sheet.createRow(0);
            for (int i = 0; i < HEADERS.length; i++) {
                header.createCell(i).setCellValue(HEADERS[i]);
            }

            // Row 1: positive / true
            writeRow(sheet, 1, dateStyle,
                    42, 9007199254740991L, 3.25, 123.5, 1000, 2012000046, 7,
                    true, 1.0, 90, 5);

            // Row 2: negative / false / 0
            writeRow(sheet, 2, dateStyle,
                    -5, -100, -2.5, 0.0, 0, 3.14, 9,
                    false, 0.0, 0, 0);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static void writeRow(final Sheet sheet, final int rowIdx, final CellStyle dateStyle,
                                 final int intVal, final long longVal, final double doubleVal,
                                 final double bigDecVal, final double bigIntVal, final double numStrVal,
                                 final double charVal, final boolean boolCell, final double boolNumVal,
                                 final double durationSeconds, final double periodDays) {

        final Row row = sheet.createRow(rowIdx);

        // date/time: the same base time as a NUMERIC (date-formatted) cell
        setDateCell(row, 0, BASE_DATE, dateStyle);
        setDateCell(row, 1, BASE_DATE, dateStyle);
        setDateCell(row, 2, BASE_DATE, dateStyle);
        setDateCell(row, 3, BASE_DATE, dateStyle);
        setDateCell(row, 4, BASE_DATE, dateStyle);
        setDateCell(row, 5, BASE_DATE, dateStyle);

        row.createCell(6).setCellValue(intVal);
        row.createCell(7).setCellValue(longVal);
        row.createCell(8).setCellValue(doubleVal);
        row.createCell(9).setCellValue(bigDecVal);
        row.createCell(10).setCellValue(bigIntVal);
        row.createCell(11).setCellValue(numStrVal);         // NUMERIC -> String
        row.createCell(12).setCellValue(charVal);           // NUMERIC -> char
        row.createCell(13).setCellValue(boolCell);          // BOOLEAN cell
        row.createCell(14).setCellValue(boolNumVal);        // NUMERIC -> Boolean
        row.createCell(15).setCellValue(durationSeconds);   // NUMERIC -> Duration(seconds)
        row.createCell(16).setCellValue(periodDays);        // NUMERIC -> Period(days)
    }

    private static List<NumericCellRow> importRows(final byte[] bytes, final PxlImportWorkbookOption option) throws Exception {
        return pxl.importExcel()
                .override(option)
                .sheet(NumericCellRow.class, Arrays.asList("Numeric"))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    // Writes an Employee header + 2 rows (Alice/Bob) into an arbitrary sheet as NUMERIC/BOOLEAN/date cells.
    private static void writeEmployeeSheet(final Sheet sheet, final CellStyle dateStyle) {
        final Row header = sheet.createRow(0);
        final String[] headers = {"Name", "Age", "Salary", "Active", "HireDate", "Grade", "Department"};
        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }

        final Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("Alice");
        r1.createCell(1).setCellValue(30);                  // NUMERIC
        r1.createCell(2).setCellValue(50000.5);             // NUMERIC -> BigDecimal
        r1.createCell(3).setCellValue(true);                // BOOLEAN cell
        final Cell h1 = r1.createCell(4);
        h1.setCellValue(dateOf(LocalDate.of(2020, 1, 15))); // NUMERIC date cell
        h1.setCellStyle(dateStyle);
        r1.createCell(5).setCellValue("A");
        r1.createCell(6).setCellValue("Engineering");

        final Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("Bob");
        r2.createCell(1).setCellValue(42);
        r2.createCell(2).setCellValue(72000);
        r2.createCell(3).setCellValue(false);
        final Cell h2 = r2.createCell(4);
        h2.setCellValue(dateOf(LocalDate.of(2018, 7, 1)));
        h2.setCellStyle(dateStyle);
        r2.createCell(5).setCellValue("B");
        r2.createCell(6).setCellValue("Sales");
    }

    // Writes a Department header + 2 rows (ENG/SAL) into an arbitrary sheet.
    private static void writeDepartmentSheet(final Sheet sheet) {
        final Row header = sheet.createRow(0);
        final String[] headers = {"Code", "DepartmentName", "Headcount"};
        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }

        final Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("ENG");
        r1.createCell(1).setCellValue("Engineering");
        r1.createCell(2).setCellValue(12);

        final Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("SAL");
        r2.createCell(1).setCellValue("Sales");
        r2.createCell(2).setCellValue(7);
    }

    private static void assertEmployees(final List<Employee> employees) {
        assertThat(employees).hasSize(2);

        final Employee alice = employees.get(0);
        assertThat(alice.getName()).isEqualTo("Alice");
        assertThat(alice.getAge()).isEqualTo(30);
        assertThat(alice.getSalary()).isEqualByComparingTo("50000.5");
        assertThat(alice.getActive()).isTrue();                                 // BOOLEAN cell
        assertThat(alice.getHireDate()).isEqualTo(LocalDate.of(2020, 1, 15));   // NUMERIC date cell
        assertThat(alice.getGrade()).isEqualTo(Grade.A);
        assertThat(alice.getDepartment()).isEqualTo("Engineering");

        final Employee bob = employees.get(1);
        assertThat(bob.getName()).isEqualTo("Bob");
        assertThat(bob.getActive()).isFalse();
        assertThat(bob.getGrade()).isEqualTo(Grade.B);
    }

    // A fixture with the Terminal/Destination columns vertically merged. (The corresponding cells in row 2 are the non-top-left part of the merge, so they are empty)
    private static byte[] buildMergedFixture() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            final Sheet sheet = workbook.createSheet("Merged");

            final Row header = sheet.createRow(0);
            final String[] headers = {"Region", "Terminal", "Stop", "Destination"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }

            final Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("North");
            r1.createCell(1).setCellValue("T1");    // merge top-left
            r1.createCell(2).setCellValue("S1");
            r1.createCell(3).setCellValue("D1");    // merge top-left

            final Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue("South");
            r2.createCell(2).setCellValue("S2");
            // Terminal(1), Destination(3) cells are not created -> non-top-left of the merge (empty cell)

            sheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));   // Terminal vertical merge
            sheet.addMergedRegion(new CellRangeAddress(1, 2, 3, 3));   // Destination vertical merge

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static List<MergedRow> importMerged(final byte[] bytes, final PxlImportWorkbookOption option) throws Exception {
        return pxl.importExcel()
                .override(option)
                .sheet(MergedRow.class, Arrays.asList("Merged"))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    private interface SheetBuilder {
        void build(Sheet sheet);
    }

    private static byte[] sheet(final String sheetName, final SheetBuilder builder) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            builder.build(workbook.createSheet(sheetName));
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    // 1 header row + N STRING data rows
    private static byte[] stringSheet(final String sheetName, final String[] headers, final String[][] dataRows) throws Exception {
        return sheet(sheetName, s -> {
            final Row header = s.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }
            for (int r = 0; r < dataRows.length; r++) {
                final Row row = s.createRow(r + 1);
                for (int c = 0; c < dataRows[r].length; c++) {
                    row.createCell(c).setCellValue(dataRows[r][c]);
                }
            }
        });
    }

    private static <T> List<T> importList(final byte[] bytes, final String sheetName, final Class<T> rowClass,
                                          final PxlImportWorkbookOption option) throws Exception {
        return pxl.importExcel()
                .override(option)
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    // ------------------------------------------------------------------
    // Resource ownership: fromStream does not close the InputStream passed by the caller
    // ------------------------------------------------------------------

    @Test
    public void importFromStream_doesNotCloseCallerStream() throws Exception {
        final byte[] bytes = stringSheet("People",
                new String[]{"Name", "Age", "Salary", "Active", "HireDate", "Grade", "Department"},
                new String[][]{{"Alice", "30", "50000.50", "true", "2020-01-15", "A", "Engineering"}});

        final boolean[] closed = {false};
        final ByteArrayInputStream tracking = new ByteArrayInputStream(bytes) {
            @Override
            public void close() {
                closed[0] = true;
            }
        };

        final List<Employee> people = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromStream(tracking);

        assertThat(people).hasSize(1);
        assertThat(closed[0]).as("fromStream must not close the caller's stream").isFalse();
    }

    // ------------------------------------------------------------------
    // Non-streaming import: NUMERIC date/number/boolean cells
    // ------------------------------------------------------------------

    @Test
    public void importExcel_numericCells_nonStreaming_parsed() throws Exception {
        final List<NumericCellRow> rows = importRows(buildNumericFixture(), null);

        assertThat(rows).hasSize(2);

        final NumericCellRow r1 = rows.get(0);
        // NUMERIC (date-formatted) cell -> date/time
        assertThat(r1.getLocalDate()).isEqualTo(LocalDate.of(2023, 6, 15));
        assertThat(r1.getLocalTime()).isEqualTo(LocalTime.of(10, 30, 45));
        assertThat(r1.getLocalDateTime()).isEqualTo(BASE);
        assertThat(r1.getJavaDate()).isEqualTo(BASE_DATE);
        assertThat(r1.getZonedDateTime()).isEqualTo(BASE.atZone(ZONE));
        assertThat(r1.getOffsetDateTime()).isEqualTo(BASE.atZone(ZONE).toOffsetDateTime());
        // NUMERIC cell -> number
        assertThat(r1.getIntVal()).isEqualTo(42);
        assertThat(r1.getLongVal()).isEqualTo(9007199254740991L);
        assertThat(r1.getDoubleVal()).isEqualTo(3.25);
        assertThat(r1.getBigDec()).isEqualByComparingTo(new BigDecimal("123.5"));
        assertThat(r1.getBigInt()).isEqualTo(BigInteger.valueOf(1000));
        // NUMERIC cell -> String (integer as-is, without exponent notation)
        assertThat(r1.getNumericAsString()).isEqualTo("2012000046");
        // NUMERIC cell -> char (first character)
        assertThat(r1.getCharFromNumeric()).isEqualTo('7');
        // BOOLEAN cell -> Boolean
        assertThat(r1.getBoolFromBoolean()).isTrue();
        // NUMERIC cell -> Boolean
        assertThat(r1.getBoolFromNumeric()).isTrue();
        // NUMERIC cell -> Duration(seconds) / Period(days)
        assertThat(r1.getDurationFromNumeric()).isEqualTo(Duration.ofSeconds(90));
        assertThat(r1.getPeriodFromNumeric()).isEqualTo(Period.ofDays(5));

        final NumericCellRow r2 = rows.get(1);
        assertThat(r2.getIntVal()).isEqualTo(-5);
        assertThat(r2.getLongVal()).isEqualTo(-100L);
        assertThat(r2.getDoubleVal()).isEqualTo(-2.5);
        assertThat(r2.getBigInt()).isEqualTo(BigInteger.ZERO);
        assertThat(r2.getBigDec()).isEqualByComparingTo(BigDecimal.ZERO);
        assertThat(r2.getNumericAsString()).isEqualTo("3.14");
        assertThat(r2.getCharFromNumeric()).isEqualTo('9');
        // false BOOLEAN cell / 0 NUMERIC -> false
        assertThat(r2.getBoolFromBoolean()).isFalse();
        assertThat(r2.getBoolFromNumeric()).isFalse();
        assertThat(r2.getDurationFromNumeric()).isEqualTo(Duration.ZERO);
        assertThat(r2.getPeriodFromNumeric()).isEqualTo(Period.ZERO);
    }

    // ------------------------------------------------------------------
    // Streaming import: the same file via the stream reader (NumberToTextConverter / StreamingCell path)
    // ------------------------------------------------------------------

    @Test
    public void importExcel_numericCells_streaming_parsed() throws Exception {
        // The stream reader does not support getFirstRowNum(), so the header/data row indices (1-based) are specified.
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<NumericCellRow> rows = importRows(buildNumericFixture(), option);

        assertThat(rows).hasSize(2);

        final NumericCellRow r1 = rows.get(0);
        assertThat(r1.getLocalDate()).isEqualTo(LocalDate.of(2023, 6, 15));
        assertThat(r1.getLocalTime()).isEqualTo(LocalTime.of(10, 30, 45));
        assertThat(r1.getLocalDateTime()).isEqualTo(BASE);
        // The remaining date/time types also decode correctly from date-formatted NUMERIC StreamingCells.
        assertThat(r1.getJavaDate()).isEqualTo(BASE_DATE);
        assertThat(r1.getZonedDateTime()).isEqualTo(BASE.atZone(ZONE));
        assertThat(r1.getOffsetDateTime()).isEqualTo(BASE.atZone(ZONE).toOffsetDateTime());
        assertThat(r1.getIntVal()).isEqualTo(42);
        assertThat(r1.getLongVal()).isEqualTo(9007199254740991L);
        assertThat(r1.getBigDec()).isEqualByComparingTo(new BigDecimal("123.5"));
        // NUMERIC -> String renders via DataFormatter (General) -> integer as-is, without exponent notation
        assertThat(r1.getNumericAsString()).isEqualTo("2012000046");
        assertThat(r1.getCharFromNumeric()).isEqualTo('7');
        assertThat(r1.getBoolFromBoolean()).isTrue();
        assertThat(r1.getBoolFromNumeric()).isTrue();
        assertThat(r1.getDurationFromNumeric()).isEqualTo(Duration.ofSeconds(90));
        assertThat(r1.getPeriodFromNumeric()).isEqualTo(Period.ofDays(5));

        final NumericCellRow r2 = rows.get(1);
        assertThat(r2.getNumericAsString()).isEqualTo("3.14");
        assertThat(r2.getBoolFromBoolean()).isFalse();
        assertThat(r2.getBoolFromNumeric()).isFalse();
    }

    // Streaming read of every date/time type from date-formatted NUMERIC cells.
    // PXL exports the 7 date/time columns without a pattern/masking (-> NUMERIC date-formatted cells), then the
    // stream reader decodes each date-formatted StreamingCell via POI's getXxxCellValue()/isCellDateFormatted path.
    // (This is the path a former StreamingCell workaround side-stepped; the workaround is no longer needed.)
    @Test
    public void importExcel_dateTimeColumns_streaming_readFromNumericDateCells() throws Exception {
        final AllTypesRow expected = Fixtures.sampleAllTypesRow();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("AllTypes", Arrays.asList(expected), AllTypesRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        // The stream reader does not support getFirstRowNum(), so the header/data row indices (1-based) are specified.
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<AllTypesRow> rows = pxl.importExcel()
                .override(option)
                .sheet(AllTypesRow.class, Arrays.asList("AllTypes"))
                .fromFile(excelFile);

        assertThat(rows).hasSize(1);
        final AllTypesRow row = rows.get(0);

        // All 7 date/time types decode correctly from a date-formatted NUMERIC StreamingCell.
        assertThat(row.getJavaDate()).isEqualTo(expected.getJavaDate());
        assertThat(row.getLocalDate()).isEqualTo(expected.getLocalDate());
        assertThat(row.getLocalTime()).isEqualTo(expected.getLocalTime());
        assertThat(row.getLocalDateTime()).isEqualTo(expected.getLocalDateTime());
        assertThat(row.getZonedDateTime()).isEqualTo(expected.getZonedDateTime());
        assertThat(row.getOffsetTime()).isEqualTo(expected.getOffsetTime());
        assertThat(row.getOffsetDateTime()).isEqualTo(expected.getOffsetDateTime());
    }

    // Builds an XLSX whose single data row holds a plain integer, a date-formatted date cell, and a large integer,
    // all bound to String columns.
    private static byte[] buildStringFromNumericFixture() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            final Sheet sheet = workbook.createSheet("StrNum");

            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));

            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("PlainInt");
            header.createCell(1).setCellValue("DateFormatted");
            header.createCell(2).setCellValue("LargeInt");

            final Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(123);                 // plain General numeric
            final Cell dateCell = data.createCell(1);
            dateCell.setCellValue(dateOf(LocalDate.of(2020, 1, 15)));
            dateCell.setCellStyle(dateStyle);                     // date-formatted numeric
            data.createCell(2).setCellValue(2012000046);          // large integer, General

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    // Reading NUMERIC cells (incl. a date-formatted one) into String via the stream reader must render them with
    // the cell's display format through DataFormatter — identical to non-streaming — now that the StreamingCell
    // /NumberToTextConverter special-case has been removed from PxlStringCodec. (excel-streaming-reader reads
    // styles by default, so the StreamingCell carries its number format.)
    @Test
    public void importExcel_numericCellsToString_streaming_rendersWithDisplayFormatLikeNonStreaming() throws Exception {
        final byte[] bytes = buildStringFromNumericFixture();

        // Non-streaming baseline: DataFormatter with the cell's display format.
        final List<StringFromNumericRow> nonStreamingRows = pxl.importExcel()
                .sheet(StringFromNumericRow.class, Arrays.asList("StrNum"))
                .fromStream(new ByteArrayInputStream(bytes));
        assertThat(nonStreamingRows).hasSize(1);
        final StringFromNumericRow nonStreaming = nonStreamingRows.get(0);

        // Streaming: the stream reader does not support getFirstRowNum(), so header/data row indices are specified.
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();
        final List<StringFromNumericRow> streamingRows = pxl.importExcel()
                .override(option)
                .sheet(StringFromNumericRow.class, Arrays.asList("StrNum"))
                .fromStream(new ByteArrayInputStream(bytes));
        assertThat(streamingRows).hasSize(1);
        final StringFromNumericRow streaming = streamingRows.get(0);

        // Streaming renders with the display format (no raw serial for the date cell), without throwing.
        assertThat(streaming.getPlainInt()).isEqualTo("123");
        assertThat(streaming.getDateFormatted()).isEqualTo("2020-01-15");
        assertThat(streaming.getLargeInt()).isEqualTo("2012000046");

        // ...and it matches the non-streaming result exactly (the removed special-case was unnecessary).
        assertThat(streaming.getPlainInt()).isEqualTo(nonStreaming.getPlainInt());
        assertThat(streaming.getDateFormatted()).isEqualTo(nonStreaming.getDateFormatted());
        assertThat(streaming.getLargeInt()).isEqualTo(nonStreaming.getLargeInt());
    }

    // ------------------------------------------------------------------
    // 1904 date windowing: the two NUMERIC branches of the LocalDateTime codec diverge
    // (date-formatted cell -> getLocalDateTimeCellValue() honors the workbook's 1904 system;
    //  plain numeric cell -> DateUtil.getLocalDateTime is fixed to the 1900 system)
    // ------------------------------------------------------------------

    // The offset between the Excel 1904 and 1900 date systems, in whole days.
    private static final int DATE_1904_OFFSET_DAYS = 1462;

    // Builds a 1904-mode XLSX whose single data row has "Formatted" (date-formatted) and "Plain" (bare number)
    // cells holding the identical raw serial.
    private static byte[] build1904Fixture(final double serial) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            // Switch the workbook to the 1904 date system (no public setter on XSSFWorkbook -> set the schema flag).
            final CTWorkbook ct = workbook.getCTWorkbook();
            final CTWorkbookPr pr = ct.isSetWorkbookPr() ? ct.getWorkbookPr() : ct.addNewWorkbookPr();
            pr.setDate1904(true);

            final Sheet sheet = workbook.createSheet("Windowing");

            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));

            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Formatted");
            header.createCell(1).setCellValue("Plain");

            final Row data = sheet.createRow(1);
            final Cell formattedCell = data.createCell(0);  // date-formatted -> isCellDateFormatted == true
            formattedCell.setCellValue(serial);
            formattedCell.setCellStyle(dateStyle);
            data.createCell(1).setCellValue(serial);        // bare number -> isCellDateFormatted == false

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    @Test
    public void importExcel_1904Workbook_dateFormattedVsPlainNumeric_windowingDiffers() throws Exception {
        // A raw serial that denotes BASE under the 1904 date system; both cells hold exactly this value.
        final double serial = DateUtil.getExcelDate(BASE, true);

        final List<DateWindowingRow> rows = pxl.importExcel()
                .sheet(DateWindowingRow.class, Arrays.asList("Windowing"))
                .fromStream(new ByteArrayInputStream(build1904Fixture(serial)));

        assertThat(rows).hasSize(1);
        final DateWindowingRow row = rows.get(0);

        // Date-formatted cell: getLocalDateTimeCellValue() honors the workbook's 1904 windowing -> the intended date.
        assertThat(row.getFormatted()).isEqualTo(BASE);
        // Plain numeric cell (same raw serial): DateUtil.getLocalDateTime is fixed to 1900 windowing -> 1462 days earlier.
        assertThat(row.getPlain()).isEqualTo(BASE.minusDays(DATE_1904_OFFSET_DAYS));
        // The crux: reading the identical raw serial, the two branches diverge by exactly the 1904 offset.
        assertThat(row.getFormatted()).isEqualTo(row.getPlain().plusDays(DATE_1904_OFFSET_DAYS));
    }

    // ------------------------------------------------------------------
    // Merged cells (importEachCellOfMergedRegion)
    // ------------------------------------------------------------------

    @Test
    public void importExcel_mergedRegion_expanded_fillsValue() throws Exception {
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importEachCellOfMergedRegion(true)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<MergedRow> rows = importMerged(buildMergedFixture(), option);

        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).getRegion()).isEqualTo("North");
        assertThat(rows.get(0).getTerminal()).isEqualTo("T1");
        assertThat(rows.get(0).getDestination()).isEqualTo("D1");
        // The merged cells are filled with the same value.
        assertThat(rows.get(1).getRegion()).isEqualTo("South");
        assertThat(rows.get(1).getTerminal()).isEqualTo("T1");
        assertThat(rows.get(1).getStop()).isEqualTo("S2");
        assertThat(rows.get(1).getDestination()).isEqualTo("D1");
    }

    @Test
    public void importExcel_mergedRegion_default_notExpanded() throws Exception {
        final List<MergedRow> rows = importMerged(buildMergedFixture(), null);

        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).getTerminal()).isEqualTo("T1");
        assertThat(rows.get(0).getDestination()).isEqualTo("D1");
        // By default, the non-top-left cells of the merge remain empty (null).
        assertThat(rows.get(1).getRegion()).isEqualTo("South");
        assertThat(rows.get(1).getTerminal()).isNull();
        assertThat(rows.get(1).getStop()).isEqualTo("S2");
        assertThat(rows.get(1).getDestination()).isNull();
    }

    @Test
    public void importExcel_mergedRegion_streaming_throws() throws Exception {
        // The streaming approach cannot fill each cell of a merged region with the same value, so an exception is thrown.
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .importEachCellOfMergedRegion(true)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final byte[] bytes = buildMergedFixture();
        assertThrows(io.github.hclimkr.pxl.exception.PxlException.class, () -> importMerged(bytes, option));
    }

    // ------------------------------------------------------------------
    // Superclass sheet inheritance / override (importOverrideSuperClassSheet)
    // ------------------------------------------------------------------

    @Test
    public void importExcel_inheritedSheet_overrideResolves() throws Exception {
        // Prepare a file with Employees/Departments sheets using PXL.
        final CompanyWorkbook source = new CompanyWorkbook();
        source.setWorkbookName("Acme");
        source.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales")));
        source.setDepartments(Arrays.asList(
                Fixtures.department("ENG", "Engineering", 12),
                Fixtures.department("SAL", "Sales", 7)));

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(source)
                .override(noValidationOption())
                .toFile(excelFile);

        final SubCompanyWorkbook sub = pxl.importExcel()
                .workbookName("Acme")
                .workbook(SubCompanyWorkbook.class)
                .fromFile(excelFile);

        // employees is not overridden -> both superclass and subclass fields are bound.
        assertThat(sub.employees).as("subclass employees").hasSize(2);
        assertThat(((SuperCompanyWorkbook) sub).employees).as("superclass employees").hasSize(2);

        // departments is overridden -> only the subclass field, the superclass field is null
        assertThat(sub.departments).as("subclass departments").hasSize(2);
        assertThat(((SuperCompanyWorkbook) sub).departments).as("superclass departments (overridden)").isNull();
    }

    // ------------------------------------------------------------------
    // Sheet name mapping via the importSheetNames option
    // ------------------------------------------------------------------

    @Test
    public void importExcel_sheetNameMapping_binds() throws Exception {
        // The actual sheet name is "Staff", which differs from the DTO's @PxlSheet(name="Employees").
        final byte[] bytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            writeEmployeeSheet(workbook.createSheet("Staff"), dateStyle);
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        }

        // Map the employees field to the actual sheet "Staff".
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .fieldName("employees")
                .importSheetNames(Arrays.asList("Staff"))
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final CompanyWorkbook workbook = pxl.importExcel()
                .workbookName("Acme")
                .override(option)
                .workbook(CompanyWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes));

        assertEmployees(workbook.getEmployees());
        // Since there is no "Departments" sheet (not required), it stays null.
        assertThat(workbook.getDepartments()).isNull();
    }

    @Test
    public void importExcel_workbookNameAndOverrideAfterWorkbookConfig_binds() throws Exception {
        // Same mapping as importExcel_sheetNameMapping_binds, but workbookName(...)/override(...) are chained
        // after workbook(...) on the source step - the chain position must not change the outcome.
        final byte[] bytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            writeEmployeeSheet(workbook.createSheet("Staff"), dateStyle);
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        }

        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .fieldName("employees")
                .importSheetNames(Arrays.asList("Staff"))
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final CompanyWorkbook workbook = pxl.importExcel()
                .workbook(CompanyWorkbook.class)
                .workbookName("Acme")
                .override(option)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(Pxl.getWorkbookNameFromWorkbookObject(workbook)).isEqualTo("Acme");
        assertEmployees(workbook.getEmployees());
    }

    @Test
    public void importExcel_candidateSheetNames_laterNameMatches_binds() throws Exception {
        // The workbook has only a "Sales" sheet. In the candidate sheet name list of a Pxl sheet-form import,
        // verify that the earlier name ("Nonexistent") does not exist and the later "Sales" matches and binds.
        final byte[] bytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            writeEmployeeSheet(workbook.createSheet("Sales"), dateStyle);
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        }

        final List<Employee> employees = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Nonexistent", "Sales"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertEmployees(employees);
    }

    // ------------------------------------------------------------------
    // Multi-sheet import in sheet form — sheet() returns a source step, so one builder is reused per sheet
    // ------------------------------------------------------------------

    @Test
    public void importExcel_sheetCalledTwicePerBuilder_bindsEachSheetIndependently() throws Exception {
        final byte[] bytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            writeEmployeeSheet(workbook.createSheet("Employees"), dateStyle);
            writeDepartmentSheet(workbook.createSheet("Departments"));
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        }

        // Unlike the export builder, sheet(...) returns a source step rather than the builder itself, so the calls
        // are not chained; the same builder instance is reused once per sheet, each with its own row class.
        final PxlExcelImportBuilder builder = pxl.importExcel();

        final List<Employee> employees = builder
                .sheet(Employee.class, Arrays.asList("Employees"))
                .fromStream(new ByteArrayInputStream(bytes));
        final List<Department> departments = builder
                .sheet(Department.class, Arrays.asList("Departments"))
                .fromStream(new ByteArrayInputStream(bytes));

        // The second call must not be affected by the first (no state leaks between source steps).
        assertEmployees(employees);
        assertThat(departments).extracting(Department::getCode).containsExactly("ENG", "SAL");
        assertThat(departments.get(0).getDepartmentName()).isEqualTo("Engineering");
        assertThat(departments.get(0).getHeadcount()).isEqualTo(12);
    }

    // ------------------------------------------------------------------
    // Reading an external XLS (HSSF) file
    // ------------------------------------------------------------------

    @Test
    public void importXls_externalHssf_reads() throws Exception {
        final byte[] bytes;
        try (Workbook workbook = new HSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            writeEmployeeSheet(workbook.createSheet("People"), dateStyle);
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        }

        final List<Employee> people = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertEmployees(people);
    }

    // ------------------------------------------------------------------
    // Inverted explicit import row indices (last < first) -> exception
    // ------------------------------------------------------------------

    @Test
    public void importExcel_rowIndexInverted_throws() throws Exception {
        final byte[] bytes = stringSheet("S", new String[]{"Name"}, new String[][]{{"Alice"}});
        assertThrows(PxlDataException.class, () -> pxl.importExcel()
                .workbookName("W")
                .workbook(InvertedRowWorkbook.class)
                .fromStream(new ByteArrayInputStream(bytes)));
    }

    // ------------------------------------------------------------------
    // An empty import data range returns an empty result without an exception
    // ------------------------------------------------------------------

    @Test
    public void importExcel_emptyDataRange_returnsEmpty() throws Exception {
        final byte[] bytes = stringSheet("S", new String[]{"Name"}, new String[][]{{"Alice"}, {"Bob"}});

        // Set the data start row far past the actual data -> empty result
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importFirstDataRowIndex(100)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        assertThat(importList(bytes, "S", Employee.class, option)).isEmpty();
    }

    // ------------------------------------------------------------------
    // Stream reader import (importUsingStreamReader) - the header/data row indices are specified.
    // ------------------------------------------------------------------

    @Test
    public void importExcel_streamReader_readsAllRows() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");
        final Employee bob = Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice, bob), Employee.class)
                .override(noValidationOption())
                .toFile(excelFile);

        // The stream reader does not support getFirstRowNum(), so the header/data row indices (1-based) are specified.
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption workbookOption = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<Employee> people = pxl.importExcel()
                .override(workbookOption)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);

        assertThat(people).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Sheet-form import — specifying the return collection type (collectionClass)
    // ------------------------------------------------------------------

    @Test
    public void importExcel_sheet_intoSet_returnsSet() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, LocalDate.of(2020, 1, 15), Grade.A, "Engineering");
        final Employee bob = Fixtures.employee("Bob", 42, "72000", false, LocalDate.of(2018, 7, 1), Grade.B, "Sales");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("People", Arrays.asList(alice, bob), Employee.class)
                .override(noValidationOption())
                .toFile(excelFile);

        // collectionClass=Set.class -> the return must be a Set implementation (default sheet() returns a List).
        @SuppressWarnings("unchecked") final Set<Employee> people =
                pxl.importExcel()
                        .sheet(Set.class, Employee.class, Arrays.asList("People"))
                        .fromFile(excelFile);

        assertThat(people).isInstanceOf(Set.class);
        assertThat(people).isNotInstanceOf(List.class);
        assertThat(people).extracting(Employee::getName).containsExactlyInAnyOrder("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Formula (exportStringAsFormula) - the computed result is read on non-streaming import.
    // ------------------------------------------------------------------

    @Test
    public void importExcel_formulaCell_evaluated() throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=2+3");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Formula", Arrays.asList(row), FormulaRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final List<FormulaRow> rows = pxl.importExcel()
                .sheet(FormulaRow.class, Arrays.asList("Formula"))
                .fromFile(excelFile);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getLabel()).isEqualTo("calc");
        // =2+3 is computed and read as "5".
        assertThat(rows.get(0).getFormula()).isEqualTo("5");
    }

    // ------------------------------------------------------------------
    // Explicit (non-default) header/data row and column indices
    // ------------------------------------------------------------------

    @Test
    public void importExcel_explicitRowAndColumnIndices_selectsRegion() throws Exception {
        // Non-default 1-based indices exercise the importer's 1-based -> 0-based bound calculations
        // (the "else" branches a default-positioned sheet skips), for both rows and columns.
        final byte[] bytes = sheet("Data", s -> {
            s.createRow(0).createCell(0).setCellValue("junk");   // 0-based row 0: before the header
            final Row header = s.createRow(1);                   // 0-based row 1: header (only cols 1..2 used)
            header.createCell(0).setCellValue("pad");
            header.createCell(1).setCellValue("Name");
            header.createCell(2).setCellValue("Age");
            header.createCell(3).setCellValue("extra");
            final String[][] data = {
                    {"pad", "Alice", "30", "extra"},   // 0-based row 2: first data row
                    {"pad", "Bob", "42", "extra"},     // 0-based row 3: last data row
                    {"pad", "Carol", "99", "extra"}};  // 0-based row 4: beyond the last data row -> excluded
            for (int r = 0; r < data.length; r++) {
                final Row row = s.createRow(r + 2);
                for (int c = 0; c < data[r].length; c++) {
                    row.createCell(c).setCellValue(data[r][c]);
                }
            }
        });

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

        final List<Employee> employees = importList(bytes, "Data", Employee.class, option);

        // only the Name/Age columns (0-based 1..2) and the rows 0-based 2..3 are read; Carol is excluded
        assertThat(employees).extracting(Employee::getName).containsExactly("Alice", "Bob");
        assertThat(employees.get(0).getAge()).isEqualTo(30);
        assertThat(employees.get(1).getAge()).isEqualTo(42);
    }

    // ------------------------------------------------------------------
    // @PxlRowIndex field-type injection (Long / Short / Byte / Number)
    // ------------------------------------------------------------------

    @Test
    public void importExcel_rowIndexFieldTypes_injectedPerType() throws Exception {
        // The 1-based spreadsheet row number is injected into the @PxlRowIndex field,
        // one branch per numeric field type. (The unsupported String type is covered in PxlExceptionTests.)
        final byte[] bytes = stringSheet("S", new String[]{"Name"}, new String[][]{{"Alice"}});
        assertThat(importList(bytes, "S", LongRowIndexRow.class, null).get(0).getRowIndex()).isEqualTo(2L);
        assertThat(importList(bytes, "S", ShortRowIndexRow.class, null).get(0).getRowIndex()).isEqualTo((short) 2);
        assertThat(importList(bytes, "S", ByteRowIndexRow.class, null).get(0).getRowIndex()).isEqualTo((byte) 2);
        assertThat(importList(bytes, "S", NumberRowIndexRow.class, null).get(0).getRowIndex().intValue()).isEqualTo(2);
    }

    @Test
    public void importExcel_rowIndex_multipleRowsShiftedLayout_injectsOneBasedRowNumber() throws Exception {
        // The @PxlRowIndex value is the 1-based spreadsheet row number of each imported row, not a
        // sequential counter: with the header shifted to 0-based row 2, the three data rows at 0-based
        // rows 3/4/5 receive exactly 4/5/6 (proving both per-row distinctness and the row-number semantics).
        final byte[] bytes = sheet("S", s -> {
            s.createRow(0).createCell(0).setCellValue("junk");    // 0-based row 0: before the header
            s.createRow(1).createCell(0).setCellValue("junk");    // 0-based row 1: before the header
            s.createRow(2).createCell(0).setCellValue("Name");    // 0-based row 2: header
            s.createRow(3).createCell(0).setCellValue("Alice");   // 0-based row 3: first data row
            s.createRow(4).createCell(0).setCellValue("Bob");     // 0-based row 4
            s.createRow(5).createCell(0).setCellValue("Carol");   // 0-based row 5: last data row
        });
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder()
                        .importHeaderRowIndex(3)          // 1-based -> 0-based 2
                        .importFirstDataRowIndex(4)       // 1-based -> 0-based 3
                        .build()))
                .build();

        final List<LongRowIndexRow> rows = importList(bytes, "S", LongRowIndexRow.class, option);

        assertThat(rows).extracting(LongRowIndexRow::getName).containsExactly("Alice", "Bob", "Carol");
        assertThat(rows).extracting(LongRowIndexRow::getRowIndex).containsExactly(4L, 5L, 6L);
    }

    // ------------------------------------------------------------------
    // Header/column resolution errors
    // ------------------------------------------------------------------

    @Test
    public void importExcel_noMatchingColumn_throws() throws Exception {
        // No header cell matches any column name -> "no header column" error.
        final byte[] bytes = stringSheet("S", new String[]{"Foo"}, new String[][]{{"1"}});
        assertThrows(PxlDataException.class, () -> importList(bytes, "S", LongRowIndexRow.class, null));
    }

    @Test
    public void importExcel_headerRowIndexBeyondData_throws() throws Exception {
        // A header row index past the end of the sheet -> "no header row" error.
        final byte[] bytes = stringSheet("S", new String[]{"Name"}, new String[][]{{"Alice"}});
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(PxlImportSheetOption.builder().importHeaderRowIndex(10).build()))
                .build();
        assertThrows(PxlDataException.class, () -> importList(bytes, "S", LongRowIndexRow.class, option));
    }
}
