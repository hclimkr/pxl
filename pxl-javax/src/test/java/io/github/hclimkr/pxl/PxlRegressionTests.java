package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
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
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static io.github.hclimkr.pxl.tcdata.TestExports.emit;
import static io.github.hclimkr.pxl.tcdata.TestExports.workbookOf;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Regression tests verifying that previously fixed issues remain fixed.
 * <p>
 * Each test corresponds to one such fix and fails to catch any side effect (bug reintroduction) caused by
 * source refactoring.
 * <p>
 * A fix has to hold on every terminal, so each test that exports is swept across {@link ExportDest}. The
 * import-only ones build their input with raw POI or a CSV string and never export at all.
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

    private <T> List<T> roundTrip(final ExportDest dest, final String sheetName, final List<T> rows, final Class<T> rowClass) throws Exception {
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(rowClass, rows, sheetName)
                .override(noValidationOption()), dest, testInfo);
        return pxl.importExcel()
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromStream(new ByteArrayInputStream(bytes));
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

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dateSeconds_defaultPattern_preserved(final ExportDest dest) throws Exception {
        final Date when = Date.from(LocalDateTime.of(2023, 6, 15, 10, 30, 45).atZone(ZoneId.systemDefault()).toInstant());

        final DateOnlyRow row = new DateOnlyRow();
        row.setWhen(when);

        final DateOnlyRow out = roundTrip(dest, "Dates", Arrays.asList(row), DateOnlyRow.class).get(0);
        // The seconds (45) must not be lost as 00
        assertThat(out.getWhen()).isEqualTo(when);
    }

    // ------------------------------------------------------------------
    // M8: trimming leading/trailing whitespace of an annotation-derived export column name
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void annotationColumnName_withWhitespace_trimmed(final ExportDest dest) throws Exception {
        final SpacedNameRow row = new SpacedNameRow();
        row.setValue("v");

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(SpacedNameRow.class, Arrays.asList(row), "Spaced")
                .override(noValidationOption()), dest, testInfo);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            final String header = workbook.getSheet("Spaced").getRow(0).getCell(0).getStringCellValue();
            // "  Padded Name  " -> trimmed at both ends -> "Padded Name" (inner whitespace preserved)
            assertThat(header).isEqualTo("Padded Name");
        }
    }

    // ------------------------------------------------------------------
    // Guard: invalid exportMasking regex -> exception at build time
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportMasking_invalidRegex_throws(final ExportDest dest) {
        final BadMaskingRow row = new BadMaskingRow();
        row.setValue("v");

        assertThrows(PxlArgumentException.class, () ->
                emit(pxl.exportExcel()
                        .sheet(BadMaskingRow.class, Arrays.asList(row), "Mask")
                        .override(noValidationOption()), dest, testInfo));
    }

    // ------------------------------------------------------------------
    // Guard: invalid date pattern -> exception at build time
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void columnPattern_invalid_throws(final ExportDest dest) {
        final BadPatternRow row = new BadPatternRow();
        row.setDate(LocalDate.of(2023, 1, 1));

        assertThrows(PxlArgumentException.class, () ->
                emit(pxl.exportExcel()
                        .sheet(BadPatternRow.class, Arrays.asList(row), "Bad")
                        .override(noValidationOption()), dest, testInfo));
    }

    // ------------------------------------------------------------------
    // L12: exportStringAsFormula fallback safely writes quote-prefixed text instead of a raw "="
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void formulaCell_invalidFallback_quotePrefixed(final ExportDest dest) throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=(");   // invalid formula -> setCellFormula fails -> quote-prefix fallback

        // Exports without exception and round-trips as literal text.
        final FormulaRow out = roundTrip(dest, "Formula", Arrays.asList(row), FormulaRow.class).get(0);
        assertThat(out.getLabel()).isEqualTo("calc");
        assertThat(out.getFormula()).isEqualTo("=(");
    }

    // ------------------------------------------------------------------
    // M9-B: a misdeclared converter fails fast at build time
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void misdeclaredConverter_failsFast(final ExportDest dest) {
        final BadConverterRow row = new BadConverterRow();
        row.setBad(new BadExportConverterObject());

        assertThrows(PxlArgumentException.class, () ->
                emit(pxl.exportExcel()
                        .sheet(BadConverterRow.class, Arrays.asList(row), "BadConv")
                        .override(noValidationOption()), dest, testInfo));
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

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportOptionItems_commaItems_dropdownPresent(final ExportDest dest) throws Exception {
        final OptionItemsCommaRow row = new OptionItemsCommaRow();
        row.setChoice("AT&T");

        // Option null -> exportDataValidation defaults to true. Comma items are handled via a hidden sheet, creating a dropdown without exception.
        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(OptionItemsCommaRow.class, Arrays.asList(row), "Opt"), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Opt");
            assertThat(sheet.getDataValidations()).isNotEmpty();
        }
    }

    // ------------------------------------------------------------------
    // L10: static helper null-argument guard (not a raw NPE)
    // ------------------------------------------------------------------

    @Test
    public void staticHelpers_nullArgs_handled() {
        assertThat(PxlWorkbookUtils.getWorkbookNameFromWorkbookObject(null)).isNull();
        assertThat(PxlExcelEngine.fromWorkbookObject(null))
                .isEqualTo(PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE);
    }
}
