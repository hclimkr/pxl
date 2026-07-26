package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlDataException;
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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Regression tests verifying that previously fixed issues remain fixed.
 * <p>
 * Each test corresponds to a "fixed" item in {@code .claude_doc/pxl-issues-and-history.md} and fails to catch
 * any side effect (bug reintroduction) caused by source refactoring.
 */
public class PxlRegressionTests {

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

    private <T> List<T> roundTrip(final String sheetName, final List<T> rows, final Class<T> rowClass) throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(sheetName, rows, rowClass)
                .override(noValidationOption())
                .toFile(excelFile);
        return pxl.importExcel()
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromFile(excelFile);
    }

    private static byte[] buildStringSheet(final String sheetName, final String[] headers, final String[][] dataRows) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet(sheetName);
            final Row header = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }
            for (int r = 0; r < dataRows.length; r++) {
                final Row row = sheet.createRow(r + 1);
                for (int c = 0; c < dataRows[r].length; c++) {
                    row.createCell(c).setCellValue(dataRows[r][c]);
                }
            }
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static InputStream csvStream(final String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.US_ASCII));
    }

    // ------------------------------------------------------------------
    // Date seconds-loss fix (2026-07-10): default-pattern java.util.Date round-trips down to seconds
    // ------------------------------------------------------------------

    @Test
    public void dateSeconds_defaultPattern_preserved() throws Exception {
        final Date when = Date.from(LocalDateTime.of(2023, 6, 15, 10, 30, 45).atZone(ZoneId.systemDefault()).toInstant());

        final DateOnlyRow row = new DateOnlyRow();
        row.setWhen(when);

        final DateOnlyRow out = roundTrip("Dates", Arrays.asList(row), DateOnlyRow.class).get(0);
        // The seconds (45) must not be lost as 00
        assertThat(out.getWhen()).isEqualTo(when);
    }

    // ------------------------------------------------------------------
    // M8: trimming leading/trailing whitespace of an annotation-derived export column name
    // ------------------------------------------------------------------

    @Test
    public void annotationColumnName_withWhitespace_trimmed() throws Exception {
        final SpacedNameRow row = new SpacedNameRow();
        row.setValue("v");

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .sheet("Spaced", Arrays.asList(row), SpacedNameRow.class)
                .override(noValidationOption())
                .toStream(outputStream);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            final String header = workbook.getSheet("Spaced").getRow(0).getCell(0).getStringCellValue();
            // "  Padded Name  " -> trimmed at both ends -> "Padded Name" (inner whitespace preserved)
            assertThat(header).isEqualTo("Padded Name");
        }
    }

    // ------------------------------------------------------------------
    // Guard: invalid exportMasking regex -> exception at build time
    // ------------------------------------------------------------------

    @Test
    public void exportMasking_invalidRegex_throws() {
        final BadMaskingRow row = new BadMaskingRow();
        row.setValue("v");

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportExcel()
                        .sheet("Mask", Arrays.asList(row), BadMaskingRow.class)
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // Guard: invalid date pattern -> exception at build time
    // ------------------------------------------------------------------

    @Test
    public void columnPattern_invalid_throws() {
        final BadPatternRow row = new BadPatternRow();
        row.setDate(java.time.LocalDate.of(2023, 1, 1));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportExcel()
                        .sheet("Bad", Arrays.asList(row), BadPatternRow.class)
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // L12: exportStringAsFormula fallback safely writes quote-prefixed text instead of a raw "="
    // ------------------------------------------------------------------

    @Test
    public void formulaCell_invalidFallback_quotePrefixed() throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=(");   // invalid formula -> setCellFormula fails -> quote-prefix fallback

        // Exports without exception and round-trips as literal text.
        final FormulaRow out = roundTrip("Formula", Arrays.asList(row), FormulaRow.class).get(0);
        assertThat(out.getLabel()).isEqualTo("calc");
        assertThat(out.getFormula()).isEqualTo("=(");
    }

    // ------------------------------------------------------------------
    // M9-B: a misdeclared converter fails fast at build time
    // ------------------------------------------------------------------

    @Test
    public void misdeclaredConverter_failsFast() {
        final BadConverterRow row = new BadConverterRow();
        row.setBad(new BadExportConverterObject());

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportExcel()
                        .sheet("BadConv", Arrays.asList(row), BadConverterRow.class)
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // M9-A: the cause of an exception thrown by an import converter is preserved
    // ------------------------------------------------------------------

    @Test
    public void importConverterThrows_causePreserved() throws Exception {
        final byte[] bytes = buildStringSheet("Throw", new String[]{"Value"}, new String[][]{{"trigger"}});

        // The message of the IllegalStateException("converter-boom") thrown by the converter must remain in the exception chain
        assertThatThrownBy(() -> pxl.importExcel()
                .sheet(ThrowingConverterRow.class, Arrays.asList("Throw"))
                .fromStream(new ByteArrayInputStream(bytes)))
                .isInstanceOf(PxlCellCodecException.class)
                .hasStackTraceContaining("converter-boom");
    }

    // ------------------------------------------------------------------
    // Guard: multiple null/blank values in an importUnique column are not misjudged as duplicates
    // ------------------------------------------------------------------

    @Test
    public void importUnique_multipleNulls_notDuplicate() throws Exception {
        // Both Code cells blank (=null); Name is filled so the row does not look empty
        final byte[] bytes = buildStringSheet("Unique",
                new String[]{"Code", "Name"},
                new String[][]{{"", "a"}, {"", "b"}});

        final List<UniqueCodeRow> rows = pxl.importExcel()
                .sheet(UniqueCodeRow.class, Arrays.asList("Unique"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).getCode()).isNull();
        assertThat(rows.get(1).getCode()).isNull();
        assertThat(rows).extracting(UniqueCodeRow::getName).containsExactly("a", "b");
    }

    // ------------------------------------------------------------------
    // Guard: fail-fast when same-named CSV streams double-bind to the same sheet
    // ------------------------------------------------------------------

    @Test
    public void csvDuplicateStreamNames_throws() {
        final String employeesCsv = "Name,Age,Salary,Active,HireDate,Grade,Department\nAlice,30,50000,yes,2020-01-15,A,Engineering\n";

        // Both streams "Employees" -> double-match to the employees field -> fail-fast
        assertThrows(PxlDataException.class, () -> pxl.importCsv()
                .workbookName("Acme")
                .workbook(CompanyWorkbook.class)
                .fromStreams(
                        Arrays.asList("Employees", "Employees"),
                        Arrays.asList(csvStream(employeesCsv), csvStream(employeesCsv)))
        );
    }

    // ------------------------------------------------------------------
    // Guard: exportOptionItems containing commas still create a dropdown without crashing
    // ------------------------------------------------------------------

    @Test
    public void exportOptionItems_commaItems_dropdownPresent() throws Exception {
        final OptionItemsCommaRow row = new OptionItemsCommaRow();
        row.setChoice("AT&T");

        // Option null -> exportDataValidation defaults to true. Comma items are handled via a hidden sheet, creating a dropdown without exception.
        final Workbook workbook = pxl.exportExcel()
                .sheet("Opt", Arrays.asList(row), OptionItemsCommaRow.class)
                .toWorkbook();
        try {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Opt");
            assertThat(sheet.getDataValidations()).isNotEmpty();
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // L10: static helper null-argument guard (not a raw NPE)
    // ------------------------------------------------------------------

    @Test
    public void staticHelpers_nullArgs_handled() {
        assertThat(Pxl.getWorkbookNameFromWorkbookObject(null)).isNull();
        assertThat(Pxl.getWorkbookFileFormatFromWorkbookObject(null))
                .isEqualTo(PxlFileFormat.XSSF);
    }
}
