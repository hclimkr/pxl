package io.github.hclimkr.pxl;

import com.github.pjfanning.xlsx.StreamingReader;
import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.constraint.PxlByteSize;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.constraint.PxlByteSizeValidator;
import io.github.hclimkr.pxl.internal.core.PxlContentsHandler;
import io.github.hclimkr.pxl.internal.core.PxlCoreExcelExporter;
import io.github.hclimkr.pxl.internal.i18n.PxlI18n;
import io.github.hclimkr.pxl.internal.meta.*;
import io.github.hclimkr.pxl.internal.support.*;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.styler.PxlVoidStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataCommaSeparatedNumericStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataTextStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataVerticalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderVerticalCenterTextStyler;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.util.*;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Proxy;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.Normalizer;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.Period;
import java.util.*;
import java.util.regex.Pattern;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Utility/library behavior tests.
 * <p>
 * The first part are reference tests that "do not call the PXL API", moved over from the removed deprecated suite.
 * (regex masking patterns, java.text.DecimalFormat, POI NumberToTextConverter, Unicode normalization, Apache Commons CSV parsing)
 * Ones that originally only printed output were turned into real tests with assertions, and the data is written mostly in English (ASCII).
 * <p>
 * The second part directly exercises the public {@code util/} helpers (PxlMiscUtils, PxlRegionUtils, PxlSheetUtils,
 * PxlRowUtils, PxlCellUtils) against a plain XSSF workbook, since those are consumer-facing entry points.
 */
public class PxlUtilityTests {

    // ------------------------------------------------------------------
    // Regex masking patterns (verifying the behavior of regexes that could be used for export masking)
    // ------------------------------------------------------------------

    // Email: mask from after the first 3 characters up to just before '@'
    private static final Pattern EMAIL_MASK = Pattern.compile("(?<=.{3}).(?=.*@)");
    // Name: keep the first and last characters, mask the middle (for 2 characters, mask the last)
    private static final Pattern NAME_MASK = Pattern.compile("(?<=.).(?=.)|(?<=^.).$");
    // Date: mask the character after '/' (skipping at most one character)
    private static final Pattern DATE_MASK = Pattern.compile("(?<=/.?).");
    // Time: mask the first character and the character after ':'
    private static final Pattern TIME_MASK = Pattern.compile("^.|(?<=:).");
    // Address: mask all digits
    private static final Pattern DIGIT_MASK = Pattern.compile("\\d");

    private static String mask(final Pattern pattern, final String input) {
        return pattern.matcher(input).replaceAll("*");
    }

    @Test
    public void masking_email_partiallyHidden() {
        assertThat(mask(EMAIL_MASK, "information@github.com")).isEqualTo("inf********@github.com");
        // If there is nothing to mask after the first 3 characters, it stays as-is
        assertThat(mask(EMAIL_MASK, "bob@github.com")).isEqualTo("bob@github.com");
    }

    @Test
    public void masking_personName_partiallyHidden() {
        assertThat(mask(NAME_MASK, "Tom Cruise")).isEqualTo("T********e");
        assertThat(mask(NAME_MASK, "John")).isEqualTo("J**n");
        assertThat(mask(NAME_MASK, "Bob")).isEqualTo("B*b");
        assertThat(mask(NAME_MASK, "Bo")).isEqualTo("B*");
        assertThat(mask(NAME_MASK, "A")).isEqualTo("A");
    }

    @Test
    public void masking_date_partiallyHidden() {
        assertThat(mask(DATE_MASK, "2020/01/01")).isEqualTo("2020/**/**");
    }

    @Test
    public void masking_time_partiallyHidden() {
        assertThat(mask(TIME_MASK, "01:02")).isEqualTo("*1:*2");
    }

    @Test
    public void masking_addressDigits_hidden() {
        assertThat(mask(DIGIT_MASK, "123 Main St, Apt 906-9031"))
                .isEqualTo("*** Main St, Apt ***-****");
    }

    // ------------------------------------------------------------------
    // java.text.DecimalFormat (number formatting behavior)
    // ------------------------------------------------------------------

    @Test
    public void decimalFormat_pattern_applied() {
        final DecimalFormatSymbols us = new DecimalFormatSymbols(Locale.US);   // ensure locale independence

        assertThat(new DecimalFormat("0", us).format(12345.6789)).isEqualTo("12346");
        assertThat(new DecimalFormat("0.0", us).format(12345.6789)).isEqualTo("12345.7");
        assertThat(new DecimalFormat("000000.00000", us).format(12345.6789)).isEqualTo("012345.67890");
        assertThat(new DecimalFormat("#,###.00", us).format(12345.6789)).isEqualTo("12,345.68");
        assertThat(new DecimalFormat("#.##%", us).format(0.1234)).isEqualTo("12.34%");
    }

    // ------------------------------------------------------------------
    // POI NumberToTextConverter (renders without scientific notation — the basis for PXL's number->string rendering)
    // ------------------------------------------------------------------

    @Test
    public void numberToText_largeValue_noScientificNotation() {
        // Even a large integer renders as-is rather than in scientific notation like "2.012E9".
        assertThat(NumberToTextConverter.toText(2012000046.0)).isEqualTo("2012000046");
        assertThat(NumberToTextConverter.toText(1.0)).isEqualTo("1");
        assertThat(NumberToTextConverter.toText(3.14)).isEqualTo("3.14");
    }

    // ------------------------------------------------------------------
    // Unicode normalization (NFC/NFD)
    // ------------------------------------------------------------------

    @Test
    public void unicodeNormalization_composedForm_applied() {
        final String nfc = "é";                                   // 'é' composed form (1 code point)
        final String nfd = Normalizer.normalize(nfc, Normalizer.Form.NFD);  // 'e' + combining accent (2 code points)

        assertThat(nfc.length()).isEqualTo(1);
        assertThat(nfd.length()).isEqualTo(2);
        assertThat(nfc).isNotEqualTo(nfd);
        // Converting NFD -> NFC yields the same string
        assertThat(Normalizer.normalize(nfd, Normalizer.Form.NFC)).isEqualTo(nfc);
        assertThat(Normalizer.isNormalized(nfc, Normalizer.Form.NFC)).isTrue();
    }

    // ------------------------------------------------------------------
    // Apache Commons CSV (the behavior of the CSV library PXL uses internally)
    // ------------------------------------------------------------------

    @Test
    public void commonsCsv_quotedFields_parsed() throws Exception {
        final String csv = "Name,Age\nAlice,30\nBob,42";

        try (CSVParser parser = CSVParser.parse(csv, CSVFormat.DEFAULT)) {
            final List<CSVRecord> records = parser.getRecords();

            assertThat(records).hasSize(3);            // 3 rows including the header row
            assertThat(records.get(0).get(0)).isEqualTo("Name");
            assertThat(records.get(1).get(0)).isEqualTo("Alice");
            assertThat(records.get(1).get(1)).isEqualTo("30");
            assertThat(records.get(2).get(0)).isEqualTo("Bob");
        }
    }

    // ==================================================================
    // Public util/ helpers exercised against a plain XSSF workbook
    // ==================================================================

    // ------------------------------------------------------------------
    // PxlMiscUtils (index/letter/reference conversions, styler-class check)
    // ------------------------------------------------------------------

    @Test
    public void miscUtils_columnIndexAndLetters_convertBothWays() {
        assertThat(PxlMiscUtils.convertColumnIndexToColumnString(0)).isEqualTo("A");
        assertThat(PxlMiscUtils.convertColumnIndexToColumnString(26)).isEqualTo("AA");
        assertThat(PxlMiscUtils.convertColumnStringToColumnIndex("A")).isEqualTo(0);
        assertThat(PxlMiscUtils.convertColumnStringToColumnIndex("AA")).isEqualTo(26);
    }

    @Test
    public void miscUtils_cellReferenceAndRange_formatted() {
        assertThat(PxlMiscUtils.convertIndexesToCellReferenceString(2, 1)).isEqualTo("B3");
        assertThat(PxlMiscUtils.convertIndexesToCellRangeAddressString(0, 0, 9, 3)).isEqualTo("A1:D10");
    }

    @Test
    public void miscUtils_cellReferenceToIndexes_parsedAndValidated() throws PxlArgumentException {
        final Pair<Integer, Integer> indexes = PxlMiscUtils.convertCellReferenceStringToIndexes("B3");
        assertThat(indexes.getLeft()).isEqualTo(2);
        assertThat(indexes.getRight()).isEqualTo(1);
        // A column-only reference has no row -> rejected.
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertCellReferenceStringToIndexes("B"));
    }

    @Test
    public void miscUtils_effectiveCellStyler_detectsConcreteStyler() {
        assertThat(PxlMiscUtils.isEffectiveCellStylerClass(null)).isFalse();
        // The PxlStyler interface is also the VOID sentinel -> not effective.
        assertThat(PxlMiscUtils.isEffectiveCellStylerClass(PxlStyler.class)).isFalse();
        assertThat(PxlMiscUtils.isEffectiveCellStylerClass(PxlDataVerticalCenterTextStyler.class)).isTrue();
    }

    // ------------------------------------------------------------------
    // PxlRegionUtils (merged-region lookup/copy/remove)
    // ------------------------------------------------------------------

    @Test
    public void regionUtils_getMergedRegion_returnsCoveringRegionOrNull() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));   // A1:C1

            assertThat(PxlRegionUtils.getMergedRegion(sheet, 0, 1)).isNotNull();
            assertThat(PxlRegionUtils.getMergedRegion(sheet, 5, 5)).isNull();
            assertThat(PxlRegionUtils.getMergedRegion(null, 0, 0)).isNull();
        }
    }

    @Test
    public void regionUtils_copyMergedRegionsInRow_replicatesAnchoredRegion() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Row srcRow = sheet.createRow(0);
            final Row dstRow = sheet.createRow(5);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));   // anchored on row 0

            PxlRegionUtils.copyMergedRegionsInRow(sheet, srcRow, dstRow);
            assertThat(PxlRegionUtils.getMergedRegion(sheet, 5, 0)).isNotNull();

            // The by-index overload replicates onto another row.
            sheet.createRow(7);
            PxlRegionUtils.copyMergedRegionsInRow(sheet, 0, 7);
            assertThat(PxlRegionUtils.getMergedRegion(sheet, 7, 1)).isNotNull();
        }
    }

    @Test
    public void regionUtils_copyMergedRegionsInRow_acrossRange_replicatesToEachRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Row srcRow = sheet.createRow(0);
            sheet.createRow(2);
            sheet.createRow(3);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));

            PxlRegionUtils.copyMergedRegionsInRow(sheet, srcRow, 2, 3);
            assertThat(PxlRegionUtils.getMergedRegion(sheet, 2, 0)).isNotNull();
            assertThat(PxlRegionUtils.getMergedRegion(sheet, 3, 0)).isNotNull();

            // An inverted range is a no-op.
            final int before = sheet.getNumMergedRegions();
            PxlRegionUtils.copyMergedRegionsInRow(sheet, srcRow, 3, 2);
            assertThat(sheet.getNumMergedRegions()).isEqualTo(before);
        }
    }

    @Test
    public void regionUtils_removeMergedRegionInRows_removesContainedRegions() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));   // within rows 2..3
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));   // outside
            assertThat(sheet.getNumMergedRegions()).isEqualTo(2);

            PxlRegionUtils.removeMergedRegionInRows(sheet, 2, 3);
            assertThat(sheet.getNumMergedRegions()).isEqualTo(1);
            assertThat(PxlRegionUtils.getMergedRegion(sheet, 5, 0)).isNotNull();

            // An inverted range is a no-op.
            PxlRegionUtils.removeMergedRegionInRows(sheet, 9, 8);
            assertThat(sheet.getNumMergedRegions()).isEqualTo(1);
        }
    }

    // ------------------------------------------------------------------
    // PxlSheetUtils (clone, print area)
    // ------------------------------------------------------------------

    @Test
    public void sheetUtils_cloneSheet_copiesPrintSetupAndName() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet srcSheet = workbook.createSheet("Src");
            srcSheet.getPrintSetup().setLandscape(true);
            srcSheet.setFitToPage(true);

            final Sheet clonedSheet = PxlSheetUtils.cloneSheet(workbook, workbook.getSheetIndex(srcSheet), "Cloned");

            assertThat(workbook.getSheetName(workbook.getSheetIndex(clonedSheet))).isEqualTo("Cloned");
            assertThat(clonedSheet.getPrintSetup().getLandscape()).isTrue();
            assertThat(clonedSheet.getFitToPage()).isTrue();
        }
    }

    @Test
    public void sheetUtils_setPrintArea_byIndexesAndByString() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");

            PxlSheetUtils.setPrintArea(sheet, 0, 0, 9, 3);
            assertThat(workbook.getPrintArea(workbook.getSheetIndex(sheet))).contains("$A$1:$D$10");

            PxlSheetUtils.setPrintArea(sheet, "A1:B2");
            assertThat(workbook.getPrintArea(workbook.getSheetIndex(sheet))).isNotNull();

            // A null sheet is a no-op.
            PxlSheetUtils.setPrintArea(null, 0, 0, 1, 1);
        }
    }

    // ------------------------------------------------------------------
    // PxlRowUtils (blank check, get/copy/remove rows)
    // ------------------------------------------------------------------

    @Test
    public void rowUtils_isBlankRow_detectsBlank() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            assertThat(PxlRowUtils.isBlankRow(null)).isTrue();

            final Row blankRow = sheet.createRow(0);
            blankRow.createCell(0).setBlank();
            assertThat(PxlRowUtils.isBlankRow(blankRow)).isTrue();

            final Row filledRow = sheet.createRow(1);
            filledRow.createCell(0).setCellValue("x");
            assertThat(PxlRowUtils.isBlankRow(filledRow)).isFalse();
        }
    }

    @Test
    public void rowUtils_getRow_createsWhenRequested() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            assertThat(PxlRowUtils.getRow(sheet, 3, false)).isNull();
            assertThat(PxlRowUtils.getRow(sheet, 3, true)).isNotNull();
            assertThat(PxlRowUtils.getRow(sheet, 3, false)).isNotNull();   // now exists
            assertThat(PxlRowUtils.getRow(null, 0, true)).isNull();
        }
    }

    @Test
    public void rowUtils_copyRow_duplicatesCellsToNewPosition() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Row srcRow = sheet.createRow(0);
            srcRow.createCell(0).setCellValue("A");
            srcRow.createCell(1).setCellValue(42);

            PxlRowUtils.copyRow(sheet, 0, 3);

            final Row dstRow = sheet.getRow(3);
            assertThat(dstRow).isNotNull();
            assertThat(dstRow.getCell(0).getStringCellValue()).isEqualTo("A");
            assertThat(dstRow.getCell(1).getNumericCellValue()).isEqualTo(42.0);

            // A missing source row is a no-op.
            PxlRowUtils.copyRow(sheet, 99, 100);
            assertThat(sheet.getRow(100)).isNull();
        }
    }

    @Test
    public void rowUtils_copyRowMultiply_replicatesAcrossRange() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("v");

            PxlRowUtils.copyRowMultiplyByCount(sheet, 0, 2, 3);   // fills rows 2, 3, 4

            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("v");
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("v");
            assertThat(sheet.getRow(4).getCell(0).getStringCellValue()).isEqualTo("v");
        }
    }

    @Test
    public void rowUtils_removeRows_deletesRangeAndShiftsUp() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("r0");
            sheet.createRow(1).createCell(0).setCellValue("r1");
            sheet.createRow(2).createCell(0).setCellValue("r2");

            PxlRowUtils.removeRowsByCount(sheet, 1, 1);   // remove row 1, shift row 2 up into row 1

            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("r0");
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("r2");

            // An inverted range is a no-op.
            PxlRowUtils.removeRowsByRange(sheet, 5, 2);
        }
    }

    // ------------------------------------------------------------------
    // PxlCellUtils (blank check, get/set values, merges, style, note, copy)
    // ------------------------------------------------------------------

    @Test
    public void cellUtils_isBlankCell_detectsBlank() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Row row = workbook.createSheet("S").createRow(0);
            assertThat(PxlCellUtils.isBlankCell(null)).isTrue();

            final Cell blankCell = row.createCell(0);
            blankCell.setBlank();
            assertThat(PxlCellUtils.isBlankCell(blankCell)).isTrue();

            final Cell stringCell = row.createCell(1);
            stringCell.setCellValue("x");
            assertThat(PxlCellUtils.isBlankCell(stringCell)).isFalse();
        }
    }

    @Test
    public void cellUtils_getCell_createsByIndexAndReference() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            assertThat(PxlCellUtils.getCell(sheet, 2, 1, false)).isNull();
            assertThat(PxlCellUtils.getCell(sheet, 2, 1, true)).isNotNull();

            // "B3" == row 2, column 1.
            final Cell byRef = PxlCellUtils.getCell(sheet, "B3", false);
            assertThat(byRef).isNotNull();
            assertThat(byRef.getRowIndex()).isEqualTo(2);
            assertThat(byRef.getColumnIndex()).isEqualTo(1);
        }
    }

    @Test
    public void cellUtils_getCellStringValue_perType() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Row row = workbook.createSheet("S").createRow(0);
            row.createCell(0).setCellValue(123);
            row.createCell(1).setCellValue("text");
            row.createCell(2).setCellValue(true);
            row.createCell(3).setCellFormula("1+2");
            final Cell blankCell = row.createCell(4);
            blankCell.setBlank();

            assertThat(PxlCellUtils.getCellStringValue(row.getCell(0))).isEqualTo("123");
            assertThat(PxlCellUtils.getCellStringValue(row.getCell(1))).isEqualTo("text");
            assertThat(PxlCellUtils.getCellStringValue(row.getCell(2))).isEqualTo("true");
            assertThat(PxlCellUtils.getCellStringValue(row.getCell(3))).isEqualTo("1+2");
            assertThat(PxlCellUtils.getCellStringValue(blankCell)).isNull();
            assertThat(PxlCellUtils.getCellStringValue(null)).isNull();
        }
    }

    @Test
    public void cellUtils_getCellStringValue_streamingCell_rendersWithDisplayFormat() throws Exception {
        // Build an XLSX with a plain integer, a date-formatted date cell, and a large integer.
        final byte[] bytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("S");
            final CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));

            final Row row = sheet.createRow(0);
            row.createCell(0).setCellValue(123);
            final Cell dateCell = row.createCell(1);
            dateCell.setCellValue(LocalDate.of(2020, 1, 15));
            dateCell.setCellStyle(dateStyle);
            row.createCell(2).setCellValue(2012000046);

            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        }

        // The streaming reader reads styles by default, so a StreamingCell carries its number format and
        // getCellStringValue renders it via DataFormatter (no NumberToTextConverter special-case), without throwing.
        try (Workbook streaming = StreamingReader.builder().open(new ByteArrayInputStream(bytes))) {
            final Row row = streaming.getSheetAt(0).iterator().next();
            assertThat(PxlCellUtils.getCellStringValue(row.getCell(0))).isEqualTo("123");
            assertThat(PxlCellUtils.getCellStringValue(row.getCell(1))).isEqualTo("2020-01-15");
            assertThat(PxlCellUtils.getCellStringValue(row.getCell(2))).isEqualTo("2012000046");
        }
    }

    @Test
    public void cellUtils_setCellValue_variousTypes() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");

            // Object overload: null blanks, Number -> numeric, other -> string.
            assertThat(PxlCellUtils.setCellValue(sheet, 0, 0, (Object) null, true).getCellType()).isEqualTo(CellType.BLANK);
            assertThat(PxlCellUtils.setCellValue(sheet, 0, 1, (Object) 5, true).getNumericCellValue()).isEqualTo(5.0);
            assertThat(PxlCellUtils.setCellValue(sheet, 0, 2, (Object) "hi", true).getStringCellValue()).isEqualTo("hi");

            // primitive double / boolean overloads.
            assertThat(PxlCellUtils.setCellValue(sheet, 1, 0, 3.5, true).getNumericCellValue()).isEqualTo(3.5);
            assertThat(PxlCellUtils.setCellValue(sheet, 1, 1, true, true).getBooleanCellValue()).isTrue();

            // Boolean wrapper: null blanks.
            assertThat(PxlCellUtils.setCellValue(sheet, 1, 2, (Boolean) null, true).getCellType()).isEqualTo(CellType.BLANK);

            // String / date / date-time / Date / Calendar overloads.
            assertThat(PxlCellUtils.setCellValue(sheet, 2, 0, "s", true).getStringCellValue()).isEqualTo("s");
            assertThat(PxlCellUtils.setCellValue(sheet, 2, 1, LocalDate.of(2020, 1, 1), true)).isNotNull();
            assertThat(PxlCellUtils.setCellValue(sheet, 2, 2, LocalDateTime.of(2020, 1, 1, 0, 0), true)).isNotNull();
            assertThat(PxlCellUtils.setCellValue(sheet, 2, 3, new Date(), true)).isNotNull();
            assertThat(PxlCellUtils.setCellValue(sheet, 2, 4, Calendar.getInstance(), true)).isNotNull();

            // Reference-based overload.
            assertThat(PxlCellUtils.setCellValue(sheet, "A5", (Object) "ref", true).getStringCellValue()).isEqualTo("ref");
        }
    }

    @Test
    public void cellUtils_setFormulaErrorBlank() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            assertThat(PxlCellUtils.setCellFormula(sheet, 0, 0, "1+2", true).getCellFormula()).isEqualTo("1+2");
            assertThat(PxlCellUtils.setCellErrorValue(sheet, 0, 1, FormulaError.DIV0.getCode(), true).getCellType()).isEqualTo(CellType.ERROR);

            final Cell filled = PxlCellUtils.setCellValue(sheet, 0, 2, "x", true);
            PxlCellUtils.setCellBlank(sheet, 0, 2, false);
            assertThat(filled.getCellType()).isEqualTo(CellType.BLANK);
        }
    }

    @Test
    public void cellUtils_getCellWithMerges_resolvesToAnchor() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("anchor");
            sheet.createRow(1);
            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));   // A1:A2, value in A1

            // (1,0) is blank but inside the merge -> resolves to the anchor cell A1.
            final Cell resolved = PxlCellUtils.getCellWithMerges(sheet, 1, 0);
            assertThat(resolved).isNotNull();
            assertThat(resolved.getStringCellValue()).isEqualTo("anchor");

            // A non-blank cell is returned as-is.
            assertThat(PxlCellUtils.getCellWithMerges(sheet, 0, 0).getStringCellValue()).isEqualTo("anchor");
        }
    }

    @Test
    public void cellUtils_copyCell_transfersValue() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Row row = sheet.createRow(0);
            final Cell srcCell = row.createCell(0);
            srcCell.setCellValue("copyme");
            final Cell dstCell = row.createCell(1);

            PxlCellUtils.copyCell(srcCell, dstCell);
            assertThat(dstCell.getStringCellValue()).isEqualTo("copyme");

            // A null source/destination is a no-op.
            PxlCellUtils.copyCell(null, dstCell);

            // Copy by index within the same sheet.
            final Cell dstCell2 = PxlCellUtils.copyCell(sheet, 0, 0, 2, 2, true);
            assertThat(dstCell2.getStringCellValue()).isEqualTo("copyme");

            // A missing source cell -> null.
            assertThat(PxlCellUtils.copyCell(sheet, 50, 50, 60, 60, true)).isNull();
        }
    }

    @Test
    public void cellUtils_cloneCellStyle_createsStyle() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Cell cell = workbook.createSheet("S").createRow(0).createCell(0);
            assertThat(PxlCellUtils.cloneCellStyle(cell)).isNotNull();
        }
    }

    @Test
    public void cellUtils_addNoteToCell_setsComment() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Cell cell = workbook.createSheet("S").createRow(0).createCell(0);

            // A null cell and a blank note are no-ops.
            PxlCellUtils.addNoteToCell(null, "x");
            PxlCellUtils.addNoteToCell(cell, "  ");
            assertThat(cell.getCellComment()).isNull();

            // A real note attaches a comment authored by PXL.
            PxlCellUtils.addNoteToCell(cell, "hello");
            assertThat(cell.getCellComment()).isNotNull();
            assertThat(cell.getCellComment().getAuthor()).isEqualTo(PxlConstants.PXL_CREATOR);
        }
    }

    // ------------------------------------------------------------------
    // PxlCollectionUtils (duplicate / uniqueness / all-same detection)
    // ------------------------------------------------------------------

    @Test
    public void collectionUtils_findDuplicates_returnsRepeatedElements() {
        // Repeated elements are collected; null elements are ignored; a null collection yields an empty set.
        assertThat(PxlCollectionUtils.findDuplicates(Arrays.asList("a", "b", "a", "c", "b")))
                .containsExactlyInAnyOrder("a", "b");
        assertThat(PxlCollectionUtils.findDuplicates(Arrays.asList("x", null, "x"))).containsExactly("x");
        assertThat(PxlCollectionUtils.findDuplicates(Collections.<String>emptyList())).isEmpty();
        assertThat(PxlCollectionUtils.findDuplicates((List<String>) null)).isEmpty();
    }

    @Test
    public void collectionUtils_findDuplicates_withMapper_returnsRepeatedKeys() {
        // Elements are projected through the mapper and the repeated keys are returned.
        assertThat(PxlCollectionUtils.findDuplicates(Arrays.asList("ab", "cd", "xyz"), String::length))
                .containsExactly(2);
        assertThat(PxlCollectionUtils.findDuplicates(Arrays.asList("a", "bb", "ccc"), String::length)).isEmpty();
    }

    @Test
    public void collectionUtils_hasDuplicates_shortCircuits() {
        assertThat(PxlCollectionUtils.hasDuplicates(Arrays.asList(1, 2, 3))).isFalse();
        assertThat(PxlCollectionUtils.hasDuplicates(Arrays.asList(1, 2, 2))).isTrue();
        assertThat(PxlCollectionUtils.hasDuplicates(Arrays.asList(1, null, null))).isFalse();   // null keys ignored
        assertThat(PxlCollectionUtils.hasDuplicates((List<Integer>) null)).isFalse();
        // Mapper overload.
        assertThat(PxlCollectionUtils.hasDuplicates(Arrays.asList("a", "bb", "c"), String::length)).isTrue();
        assertThat(PxlCollectionUtils.hasDuplicates(Arrays.asList("a", "bb", "ccc"), String::length)).isFalse();
    }

    @Test
    public void collectionUtils_hasAllUnique_treatsNullAsUnique() {
        assertThat(PxlCollectionUtils.hasAllUnique(Arrays.asList(1, 2, 3))).isTrue();
        assertThat(PxlCollectionUtils.hasAllUnique(Arrays.asList(1, 1))).isFalse();
        assertThat(PxlCollectionUtils.hasAllUnique((List<Integer>) null)).isTrue();
        // Mapper overload.
        assertThat(PxlCollectionUtils.hasAllUnique(Arrays.asList("a", "bb"), String::length)).isTrue();
        assertThat(PxlCollectionUtils.hasAllUnique(Arrays.asList("a", "b"), String::length)).isFalse();
    }

    @Test
    public void collectionUtils_hasAllSame_detectsUniformList() {
        assertThat(PxlCollectionUtils.hasAllSame(Arrays.asList(5, 5, 5))).isTrue();
        assertThat(PxlCollectionUtils.hasAllSame(Arrays.asList(5, 6))).isFalse();
        assertThat(PxlCollectionUtils.hasAllSame(Collections.<Integer>emptyList())).isTrue();
        assertThat(PxlCollectionUtils.hasAllSame((List<Integer>) null)).isTrue();
    }

    // ==================================================================
    // internal/support helpers (reflection, class, number, date-time, temporal-amount, workbook, type-reference)
    // ==================================================================

    // ------------------------------------------------------------------
    // PxlReflectionSupport
    // ------------------------------------------------------------------

    @Test
    public void reflectionSupport_withAnnotationValue_overridesSingleMember() throws Exception {
        final PxlColumn original = AllTypesRow.class.getDeclaredField("text").getAnnotation(PxlColumn.class);
        assertThat(original.name()).isEqualTo(new String[]{"Text"});

        final PxlColumn overridden = PxlReflectionSupport.withAnnotationValue(original, "name", new String[]{"Renamed"});
        assertThat(overridden.name()).isEqualTo(new String[]{"Renamed"});           // overridden member
        assertThat(overridden.exportSample()).isEqualTo(original.exportSample());   // other members delegate to the original
        assertThat(overridden.annotationType()).isEqualTo(PxlColumn.class);         // annotationType() delegates too
    }

    @Test
    public void reflectionSupport_withAnnotationValue_invalidKeyOrValue_throws() throws Exception {
        final PxlColumn annotation = AllTypesRow.class.getDeclaredField("text").getAnnotation(PxlColumn.class);
        // no such member
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.withAnnotationValue(annotation, "noSuchMember", "x"));
        // null new value
        assertThrows(PxlArgumentException.class, () -> PxlReflectionSupport.withAnnotationValue(annotation, "name", null));
        // incompatible value type (name() returns String, given an Integer)
        assertThrows(PxlArgumentException.class, () -> PxlReflectionSupport.withAnnotationValue(annotation, "name", 123));
    }

    @Test
    public void reflectionSupport_stringTypeConstructor_detectedAndReturned() {
        assertThat(PxlReflectionSupport.hasStringTypeConstructor(Point.class)).isTrue();   // Point(String) exists
        assertThat(PxlReflectionSupport.hasStringTypeConstructor(Object.class)).isFalse();
        assertThat(PxlReflectionSupport.getStringTypeConstructor(Point.class)).isNotNull();
        assertThat(PxlReflectionSupport.getStringTypeConstructor(Object.class)).isNull();
    }

    @Test
    public void reflectionSupport_newClassInstance_errorPaths_throw() throws Exception {
        // no no-arg constructor -> ReflectiveOperationException path
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.newClassInstance(NoDefaultCtorRow.class));
        // constructor itself throws -> InvocationTargetException path (cause propagated)
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.newClassInstance(ThrowingCtorRow.class));
        // happy path
        assertThat(PxlReflectionSupport.newClassInstance(Point.class)).isInstanceOf(Point.class);
    }

    @Test
    public void reflectionSupport_getParameterizedArgument0_variants() throws Exception {
        final Field stringList = AllTypesRow.class.getDeclaredField("stringList");   // List<String>
        assertThat(PxlReflectionSupport.getParameterizedArgument0(stringList)).isEqualTo(String.class);

        // a non-parameterized (raw) field is rejected
        final Field text = AllTypesRow.class.getDeclaredField("text");   // String
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.getParameterizedArgument0(text));

        // a nested generic (List<List<String>>) whose first argument is not a concrete Class is rejected
        final Field nested = NestedCollectionRow.class.getDeclaredField("nested");
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.getParameterizedArgument0(nested));
    }

    @Test
    public void reflectionSupport_fieldValue_getSet_andErrors() throws Exception {
        final Field textField = AllTypesRow.class.getDeclaredField("text");
        final AllTypesRow row = new AllTypesRow();

        PxlReflectionSupport.setFieldValue(textField, row, "hello");
        assertThat(PxlReflectionSupport.getFieldValue(textField, row)).isEqualTo("hello");

        // incompatible value type -> wrapped as PxlReflectionException
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.setFieldValue(textField, row, 123));
        // wrong target object type -> wrapped as PxlReflectionException
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.getFieldValue(textField, "not a row"));
    }

    @Test
    public void reflectionSupport_getToStringMethod_findsOverride() {
        assertThat(PxlReflectionSupport.getToStringMethod(Point.class)).isNotNull();   // Point overrides toString
        assertThat(PxlReflectionSupport.getToStringMethod(Object.class)).isNull();     // only Object's default
    }

    // ------------------------------------------------------------------
    // PxlClassSupport (concrete collection resolution)
    // ------------------------------------------------------------------

    @Test
    public void classSupport_getConcreteCollectionClass_resolvesInterfacesAndConcrete() throws Exception {
        assertThat(PxlClassSupport.getConcreteCollectionClass(List.class)).isEqualTo(ArrayList.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Collection.class)).isEqualTo(ArrayList.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Set.class)).isEqualTo(HashSet.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(SortedSet.class)).isEqualTo(TreeSet.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(NavigableSet.class)).isEqualTo(TreeSet.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Deque.class)).isEqualTo(ArrayDeque.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Queue.class)).isEqualTo(LinkedList.class);
        // a concrete Collection class is returned as-is
        assertThat(PxlClassSupport.getConcreteCollectionClass(LinkedList.class)).isEqualTo(LinkedList.class);
    }

    @Test
    public void classSupport_getConcreteCollectionClass_unsupported_throws() {
        // a non-collection interface
        assertThrows(PxlReflectionException.class, () -> PxlClassSupport.getConcreteCollectionClass(Comparable.class));
        // a concrete non-collection class
        assertThrows(PxlReflectionException.class, () -> PxlClassSupport.getConcreteCollectionClass(String.class));
    }

    // ------------------------------------------------------------------
    // PxlNumberSupport (range / finiteness guards)
    // ------------------------------------------------------------------

    @Test
    public void numberSupport_requireWithinRange_variants() throws Exception {
        // BigDecimal overload
        assertThat(PxlNumberSupport.requireWithinRange(new BigDecimal("5"), 0, 10, "X")).isEqualByComparingTo("5");
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(new BigDecimal("-1"), 0, 10, "X"));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(new BigDecimal("11"), 0, 10, "X"));
        // Number overload
        assertThat(PxlNumberSupport.requireWithinRange((Number) 7, 0, 10, "X")).isEqualByComparingTo("7");
        // double overload
        assertThat(PxlNumberSupport.requireWithinRange(3.0, 0, 10, "X")).isEqualByComparingTo("3");
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(100.0, 0, 10, "X"));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(Double.NaN, 0, 10, "X"));
    }

    @Test
    public void numberSupport_requireFiniteForExport_rejectsNonFinite() throws Exception {
        // finite and null values pass
        PxlNumberSupport.requireFiniteForExport((Float) 1.5F);
        PxlNumberSupport.requireFiniteForExport((Float) null);
        PxlNumberSupport.requireFiniteForExport((Double) 2.5);
        PxlNumberSupport.requireFiniteForExport((Double) null);
        // NaN / Infinity are rejected
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Float.NaN));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Float.POSITIVE_INFINITY));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Double.NaN));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Double.NEGATIVE_INFINITY));
    }

    // ------------------------------------------------------------------
    // PxlDateTimeSupport
    // ------------------------------------------------------------------

    @Test
    public void dateTimeSupport_javaDateLocalDateTime_roundTripAndNull() {
        assertThat(PxlDateTimeSupport.localDateTimeToJavaDate(null)).isNull();
        assertThat(PxlDateTimeSupport.javaDateToLocalDateTime(null)).isNull();

        final LocalDateTime ldt = LocalDateTime.of(2023, 6, 15, 10, 30, 45);
        final Date date = PxlDateTimeSupport.localDateTimeToJavaDate(ldt);
        assertThat(PxlDateTimeSupport.javaDateToLocalDateTime(date)).isEqualTo(ldt);
    }

    @Test
    public void dateTimeSupport_toExcelDateFormat_lowercasesOutsideLiterals() {
        assertThat(PxlDateTimeSupport.toExcelDateFormat("yyyy-MM-dd HH:mm:ss")).isEqualTo("yyyy-mm-dd hh:mm:ss");
        // text inside single quotes is copied verbatim (not lower-cased)
        assertThat(PxlDateTimeSupport.toExcelDateFormat("yyyy'T'HH")).isEqualTo("yyyy'T'hh");
    }

    // ------------------------------------------------------------------
    // PxlTemporalAmountSupport (DurationFormatUtils-style pattern parser)
    // ------------------------------------------------------------------

    @Test
    public void temporalAmountSupport_compileAndParse_durationAndPeriod() {
        final PxlTemporalAmountSupport.CompiledTemporalPattern durationPattern =
                PxlTemporalAmountSupport.compileTemporalPattern("HH:mm:ss");
        assertThat(PxlTemporalAmountSupport.parseDurationByPattern("01:02:03", durationPattern))
                .isEqualTo(Duration.ofHours(1).plusMinutes(2).plusSeconds(3));

        final PxlTemporalAmountSupport.CompiledTemporalPattern periodPattern =
                PxlTemporalAmountSupport.compileTemporalPattern("y'y'M'm'd'd'");
        assertThat(PxlTemporalAmountSupport.parsePeriodByPattern("1y2m3d", periodPattern))
                .isEqualTo(Period.of(1, 2, 3));
    }

    @Test
    public void temporalAmountSupport_compile_literalsAndEscapedQuote() {
        // "d''H" is: token d, an escaped literal single quote (''), token H.
        final PxlTemporalAmountSupport.CompiledTemporalPattern pattern =
                PxlTemporalAmountSupport.compileTemporalPattern("d''H");
        assertThat(PxlTemporalAmountSupport.parseDurationByPattern("5'7", pattern))
                .isEqualTo(Duration.ofDays(5).plusHours(7));
    }

    @Test
    public void temporalAmountSupport_parse_mismatchAndOverflow_throw() {
        final PxlTemporalAmountSupport.CompiledTemporalPattern hms =
                PxlTemporalAmountSupport.compileTemporalPattern("HH:mm:ss");
        // the value does not match the pattern
        assertThrows(IllegalArgumentException.class, () -> PxlTemporalAmountSupport.parseDurationByPattern("nope", hms));

        final PxlTemporalAmountSupport.CompiledTemporalPattern days = PxlTemporalAmountSupport.compileTemporalPattern("d");
        // a Period field overflows the int range (Math.toIntExact)
        assertThrows(IllegalArgumentException.class, () -> PxlTemporalAmountSupport.parsePeriodByPattern("99999999999", days));
        // a Duration overflows its range
        assertThrows(IllegalArgumentException.class,
                () -> PxlTemporalAmountSupport.parseDurationByPattern(String.valueOf(Long.MAX_VALUE), days));
    }

    // ------------------------------------------------------------------
    // PxlWorkbookSupport (public helpers)
    // ------------------------------------------------------------------

    @Test
    public void workbookSupport_makeUniqueSafeSheetName_avoidsCollision() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Data");
            assertThat(PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, "Fresh")).isEqualTo("Fresh");
            // a colliding name gets a " (2)" suffix (POI matches case-insensitively)
            assertThat(PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, "Data")).isEqualTo("Data (2)");
            workbook.createSheet("Data (2)");
            assertThat(PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, "Data")).isEqualTo("Data (3)");
        }
    }

    @Test
    public void workbookSupport_makeUniqueDefinedName_avoidsCollision() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            assertThat(PxlWorkbookSupport.makeUniqueDefinedName(workbook, "MyName")).isEqualTo("MyName");

            final Name name = workbook.createName();
            name.setNameName("MyName");
            name.setRefersToFormula("Sheet1!$A$1");

            assertThat(PxlWorkbookSupport.makeUniqueDefinedName(workbook, "MyName")).isEqualTo("MyName_2");
        }
    }

    @Test
    public void workbookSupport_createWorkbook_perFormat() throws Exception {
        assertThat(PxlWorkbookSupport.createWorkbook(PxlFileFormat.XSSF, 100)).isInstanceOf(XSSFWorkbook.class);
        assertThat(PxlWorkbookSupport.createWorkbook(PxlFileFormat.HSSF, 100)).isInstanceOf(HSSFWorkbook.class);
        assertThat(PxlWorkbookSupport.createWorkbook(PxlFileFormat.SXSSF, 100)).isInstanceOf(SXSSFWorkbook.class);
    }

    @Test
    public void workbookSupport_validateWorkbookNameFieldType_rejectsNonString() throws Exception {
        // null and a String-typed name field are accepted
        PxlWorkbookSupport.validateWorkbookNameFieldType(null);
        PxlWorkbookSupport.validateWorkbookNameFieldType(AllTypesWorkbook.class);   // String workbookName -> ok
        // a non-String @PxlWorkbookName field is rejected
        assertThrows(PxlDataException.class, () -> PxlWorkbookSupport.validateWorkbookNameFieldType(BadWorkbookNameWorkbook.class));
    }

    // ------------------------------------------------------------------
    // PxlTypeReference (super type token)
    // ------------------------------------------------------------------

    @Test
    @SuppressWarnings("deprecation")
    public void typeReference_capturesGenericArgument() {
        final PxlTypeReference<List<String>> reference = new PxlTypeReference<List<String>>() {
        };
        assertThat(reference.getType()).isInstanceOf(ParameterizedType.class);
        assertThat(((ParameterizedType) reference.getType()).getRawType()).isEqualTo(List.class);
    }

    @Test
    @SuppressWarnings({"deprecation", "rawtypes"})
    public void typeReference_rawInstantiation_throws() {
        // a raw (non-parameterized) subclass has no captured type argument -> rejected
        assertThrows(IllegalArgumentException.class, () -> new PxlTypeReference() {
        });
    }

    // ==================================================================
    // internal/meta helpers (custom-converter resolution, factory chain, toString)
    // ==================================================================

    @Test
    public void exportConverterMeta_of_resolvesAndValidates() throws Exception {
        // instance converter (Money) and toString-based (Point) both resolve
        assertThat(PxlExportColumnMeta.PxlExportConverterMeta.of(Money.class)).isNotNull();
        assertThat(PxlExportColumnMeta.PxlExportConverterMeta.of(Point.class)).isNotNull();
        // neither a converter nor a toString override -> unsupported
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(Object.class));
    }

    @Test
    public void exportConverterMeta_of_invalidSignatures_throw() {
        // return type is not String
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(BadExportConverterReturnObject.class));
        // static converter with 0 args (must take exactly the object)
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(BadExportConverterObject.class));
        // instance converter that takes an argument (must take none)
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(BadExportConverterInstanceParamObject.class));
    }

    @Test
    public void importConverterMeta_of_resolvesAndValidates() throws Exception {
        // static converter (Money) and String-constructor-based (Point) both resolve
        assertThat(PxlImportColumnMeta.PxlImportConverterMeta.of(Money.class)).isNotNull();
        assertThat(PxlImportColumnMeta.PxlImportConverterMeta.of(Point.class)).isNotNull();
        // an enum resolves even without a converter or a String constructor
        assertThat(PxlImportColumnMeta.PxlImportConverterMeta.of(Grade.class)).isNotNull();
        // neither a converter, a String constructor, nor an enum -> unsupported
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(Object.class));
    }

    @Test
    public void importConverterMeta_of_invalidSignatures_throw() {
        // return type is not the object type
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(BadImportConverterReturnObject.class));
        // converter is an instance method (must be static)
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(BadImportConverterInstanceObject.class));
        // converter does not take a single String argument
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(BadImportConverterParamObject.class));
    }

    @Test
    public void exportMeta_factoryChain_andToString() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMeta(AllTypesWorkbook.class, null);
        final List<PxlExportSheetMeta> sheetMetas =
                PxlExportSheetMeta.makeExportSheetMetas(AllTypesWorkbook.class, workbookMeta, null, false);
        assertThat(sheetMetas).isNotEmpty();

        final PxlExportSheetMeta sheetMeta = sheetMetas.get(0);
        assertThat(sheetMeta.toString()).isNotBlank();   // PxlExportSheetMeta.toString

        final List<PxlExportColumnMeta> columnMetas = PxlExportColumnMeta.makeExportColumnMetas(sheetMeta, false);
        assertThat(columnMetas).isNotEmpty();
        assertThat(columnMetas.get(0).toString()).isNotBlank();   // PxlExportColumnMeta.toString
    }

    @Test
    public void importMeta_factoryChain_andToString() throws Exception {
        final PxlImportWorkbookMeta workbookMeta =
                PxlImportWorkbookMeta.makeImportWorkbookMeta(AllTypesWorkbook.class, null);
        // instance sheet-option accessor: an empty option list yields null at any index
        assertThat(workbookMeta.getImportSheetOption(0)).isNull();

        final List<PxlImportSheetMeta> sheetMetas =
                PxlImportSheetMeta.makeImportSheetMetas(AllTypesWorkbook.class, workbookMeta, null);
        assertThat(sheetMetas).isNotEmpty();

        final PxlImportSheetMeta sheetMeta = sheetMetas.get(0);
        assertThat(sheetMeta.toString()).isNotBlank();   // PxlImportSheetMeta.toString

        final List<PxlImportColumnMeta> columnMetas = PxlImportColumnMeta.makeImportColumnMetas(sheetMeta);
        assertThat(columnMetas).isNotEmpty();
        assertThat(columnMetas.get(0).toString()).isNotBlank();   // PxlImportColumnMeta.toString
    }

    // ==================================================================
    // internal/core helpers
    // ==================================================================

    @Test
    @SuppressWarnings({"deprecation", "unchecked"})
    public void contentsHandler_accumulatesHeaderAndPaddedDataRows() throws Exception {
        // PxlContentsHandler is a POI event-model SheetContentsHandler (deprecated/reference-only).
        // Drive the SAX-style callbacks directly: row 0 is the header, later rows are padded to header width,
        // and skipped cells within a row are padded with empty strings.
        final PxlContentsHandler handler = new PxlContentsHandler();

        handler.startRow(0);
        handler.cell("A1", "H1", null);
        handler.cell("C1", "H3", null);   // skips B1 -> padded
        handler.endRow(0);

        handler.startRow(1);
        handler.cell("A2", "v1", null);   // shorter than the header -> padded on endRow
        handler.endRow(1);

        handler.headerFooter("ignored", true, "oddHeader");   // no-op
        handler.endSheet();                                    // no-op

        // The class exposes no getters, so read the accumulated state reflectively.
        final List<String> header = (List<String>) PxlReflectionSupport.getFieldValue(
                PxlContentsHandler.class.getDeclaredField("header"), handler);
        assertThat(header).containsExactly("H1", "", "H3");

        final List<List<String>> rows = (List<List<String>>) PxlReflectionSupport.getFieldValue(
                PxlContentsHandler.class.getDeclaredField("rows"), handler);
        assertThat(rows).hasSize(1);
        assertThat(rows.get(0)).containsExactly("v1", "", "");   // padded to the header width (3)
    }

    @Test
    public void excelExporter_singleSheetBuildWorkbook_writesSheet() throws Exception {
        // The single-sheet String overload of buildWorkbook is not used by the fluent builder (which always
        // uses the multi-sheet list overload), so exercise it directly.
        final PxlExportWorkbookMeta workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null);
        final List<Employee> rows = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, null, Grade.B, "Sales"));

        final Workbook workbook = PxlCoreExcelExporter.buildWorkbook("People", rows, Employee.class, workbookMeta, null);

        final Sheet sheet = workbook.getSheet("People");
        assertThat(sheet).isNotNull();
        assertThat(sheet.getRow(0)).as("header row").isNotNull();
        assertThat(sheet.getLastRowNum()).isEqualTo(2);   // header + 2 data rows (0-based last index 2)
    }

    @Test
    public void excelExporter_multiSheetBuildWorkbook_validatesArguments() throws Exception {
        final List<Employee> employees = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"));
        final List<Collection<?>> objects = Arrays.<Collection<?>>asList(employees);
        final List<Class<?>> classes = Arrays.<Class<?>>asList(Employee.class);

        // sheetNames vs sheetObjects count mismatch
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1", "S2"), objects, classes, PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
        // sheetNames vs rowClasses count mismatch
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1"), objects, Arrays.<Class<?>>asList(Employee.class, Employee.class),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
        // a null sheet-object element
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1"), Arrays.<Collection<?>>asList((Collection<?>) null), classes,
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
        // a null row-class element
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1"), objects, Arrays.<Class<?>>asList((Class<?>) null),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
    }

    @Test
    public void excelExporter_multiSheetBuildSampleWorkbook_validatesArguments() throws Exception {
        final List<Class<?>> classes = Arrays.<Class<?>>asList(Employee.class);

        // sheetNames vs rowClasses count mismatch
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildSampleWorkbook(
                Arrays.asList("S1", "S2"), classes, PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null)));
        // duplicate sheet name
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildSampleWorkbook(
                Arrays.asList("Dup", "Dup"), Arrays.<Class<?>>asList(Employee.class, Employee.class),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null)));
        // a null row-class element
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildSampleWorkbook(
                Arrays.asList("S1"), Arrays.<Class<?>>asList((Class<?>) null),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null)));
    }

    // ==================================================================
    // internal/i18n (shared UTF-8 resource-bundle loader)
    // ==================================================================

    @Test
    public void i18n_getBundleAndMessage_resolveDisableAndFallback() throws Exception {
        // i18n disabled (blank base name or null locale) -> null bundle
        assertThat(PxlI18n.getBundle("", Locale.ENGLISH)).isNull();
        assertThat(PxlI18n.getBundle("messages", (Locale) null)).isNull();
        assertThat(PxlI18n.getBundle("")).isNull();

        // a present bundle resolves; a known key returns its (locale-specific) message
        assertThat(PxlI18n.getBundle("messages", Locale.ENGLISH)).isNotNull();
        assertThat(PxlI18n.getMessage("messages", Locale.ENGLISH, "role")).isEqualTo("Role");
        assertThat(PxlI18n.getMessage("messages", Locale.KOREAN, "role")).isEqualTo("역할");
        // a missing key falls back to the key itself
        assertThat(PxlI18n.getMessage("messages", Locale.ENGLISH, "no.such.key")).isEqualTo("no.such.key");
        // i18n disabled -> the key is returned unchanged
        assertThat(PxlI18n.getMessage("", null, "role")).isEqualTo("role");
        // params overload (the message has no placeholder, so it is returned as-is)
        assertThat(PxlI18n.getMessage("messages", Locale.ENGLISH, "people", new Object[]{"x"})).isEqualTo("Staff");

        // JVM-default-locale overload resolves the bundle
        assertThat(PxlI18n.getBundle("messages")).isNotNull();

        // a missing bundle -> PxlI18nException
        assertThrows(PxlI18nException.class, () -> PxlI18n.getBundle("no-such-bundle", Locale.ENGLISH));
    }

    // ==================================================================
    // internal/constraint (PxlByteSize validator initialization)
    // ==================================================================

    @Test
    public void byteSizeValidator_initialize_rejectsInvalidBounds() throws Exception {
        final PxlByteSize annotation = ByteSizeRow.class.getDeclaredField("code").getAnnotation(PxlByteSize.class);
        final PxlByteSizeValidator validator = new PxlByteSizeValidator();

        // a valid annotation (min=0, max=5) initializes without error
        validator.initialize(annotation);

        // min < 0
        assertThrows(IllegalArgumentException.class,
                () -> validator.initialize(PxlReflectionSupport.withAnnotationValue(annotation, "min", -1)));
        // max < 0
        assertThrows(IllegalArgumentException.class,
                () -> validator.initialize(PxlReflectionSupport.withAnnotationValue(annotation, "max", -1)));
        // max < min
        assertThrows(IllegalArgumentException.class,
                () -> validator.initialize(PxlReflectionSupport.withAnnotationValue(
                        PxlReflectionSupport.withAnnotationValue(annotation, "min", 10), "max", 3)));
    }

    // ==================================================================
    // styler/ apply methods (optional stylers that no built-in default uses)
    // ==================================================================

    @Test
    public void stylers_applyDirectly_mutateStyle() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            // void styler: no-op, applies no font
            assertThat(new PxlVoidStyler().apply(workbook, workbook.createCellStyle())).isNull();

            // data text styler forces the "@" text format
            final CellStyle textStyle = workbook.createCellStyle();
            new PxlDataTextStyler().apply(workbook, textStyle);
            assertThat(textStyle.getDataFormatString()).isEqualTo("@");

            // data comma styler forces the "#,##0" numeric format
            final CellStyle commaStyle = workbook.createCellStyle();
            new PxlDataCommaSeparatedNumericStyler().apply(workbook, commaStyle);
            assertThat(commaStyle.getDataFormatString()).isEqualTo("#,##0");

            // header vertical-center styler centers vertically (and inherits a bold font from the base header styler)
            final CellStyle verticalCenterStyle = workbook.createCellStyle();
            final Font verticalCenterFont = new PxlHeaderVerticalCenterTextStyler().apply(workbook, verticalCenterStyle);
            assertThat(verticalCenterStyle.getVerticalAlignment()).isEqualTo(VerticalAlignment.CENTER);
            assertThat(verticalCenterFont).isNotNull();
        }
    }

    // ------------------------------------------------------------------
    // PxlCellUtils: cell-reference-based setter overloads (per value type)
    // ------------------------------------------------------------------

    @Test
    public void cellUtils_setCellValue_byReference_perType() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");

            assertThat(PxlCellUtils.setCellValue(sheet, "A1", (Object) 5, true).getNumericCellValue()).isEqualTo(5.0);
            assertThat(PxlCellUtils.setCellValue(sheet, "A2", 3.5, true).getNumericCellValue()).isEqualTo(3.5);
            assertThat(PxlCellUtils.setCellValue(sheet, "A3", true, true).getBooleanCellValue()).isTrue();
            assertThat(PxlCellUtils.setCellValue(sheet, "A4", (Boolean) Boolean.FALSE, true).getBooleanCellValue()).isFalse();
            assertThat(PxlCellUtils.setCellValue(sheet, "A5", "text", true).getStringCellValue()).isEqualTo("text");
            assertThat(PxlCellUtils.setCellValue(sheet, "A6", new Date(), true)).isNotNull();
            assertThat(PxlCellUtils.setCellValue(sheet, "A7", LocalDate.of(2020, 1, 1), true)).isNotNull();
            assertThat(PxlCellUtils.setCellValue(sheet, "A8", LocalDateTime.of(2020, 1, 1, 0, 0), true)).isNotNull();
            assertThat(PxlCellUtils.setCellValue(sheet, "A9", Calendar.getInstance(), true)).isNotNull();
            assertThat(PxlCellUtils.setCellValue(sheet, "A10", workbook.getCreationHelper().createRichTextString("rich"), true)).isNotNull();

            // formula / error / blank by reference
            assertThat(PxlCellUtils.setCellFormula(sheet, "B1", "1+2", true).getCellFormula()).isEqualTo("1+2");
            assertThat(PxlCellUtils.setCellErrorValue(sheet, "B2", FormulaError.DIV0.getCode(), true).getCellType()).isEqualTo(CellType.ERROR);
            final Cell filled = PxlCellUtils.setCellValue(sheet, "B3", "x", true);
            PxlCellUtils.setCellBlank(sheet, "B3", false);
            assertThat(filled.getCellType()).isEqualTo(CellType.BLANK);
        }
    }

    @Test
    public void cellUtils_copyCell_perCellType() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Row row = workbook.createSheet("S").createRow(0);

            // numeric
            row.createCell(0).setCellValue(42.0);
            PxlCellUtils.copyCell(row.getCell(0), row.createCell(1));
            assertThat(row.getCell(1).getNumericCellValue()).isEqualTo(42.0);

            // boolean
            row.createCell(2).setCellValue(true);
            PxlCellUtils.copyCell(row.getCell(2), row.createCell(3));
            assertThat(row.getCell(3).getBooleanCellValue()).isTrue();

            // blank
            row.createCell(4).setBlank();
            final Cell blankDst = row.createCell(5);
            blankDst.setCellValue("x");
            PxlCellUtils.copyCell(row.getCell(4), blankDst);
            assertThat(blankDst.getCellType()).isEqualTo(CellType.BLANK);

            // formula
            row.createCell(6).setCellFormula("1+2");
            PxlCellUtils.copyCell(row.getCell(6), row.createCell(7));
            assertThat(row.getCell(7).getCellFormula()).isEqualTo("1+2");

            // error
            row.createCell(8).setCellErrorValue(FormulaError.DIV0.getCode());
            PxlCellUtils.copyCell(row.getCell(8), row.createCell(9));
            assertThat(row.getCell(9).getCellType()).isEqualTo(CellType.ERROR);

            // cross-workbook copy clones the style into the destination workbook
            try (XSSFWorkbook workbook2 = new XSSFWorkbook()) {
                final Cell dst = workbook2.createSheet("T").createRow(0).createCell(0);
                PxlCellUtils.copyCell(row.getCell(0), dst);
                assertThat(dst.getNumericCellValue()).isEqualTo(42.0);
            }
        }
    }

    @Test
    public void rowUtils_copyRowMultiplyByRange_replicatesToRange() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("v");

            PxlRowUtils.copyRowMultiplyByRange(sheet, 0, 2, 4);   // copy row 0 into rows 2..4

            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("v");
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("v");
            assertThat(sheet.getRow(4).getCell(0).getStringCellValue()).isEqualTo("v");

            // an inverted range is a no-op
            PxlRowUtils.copyRowMultiplyByRange(sheet, 0, 9, 8);
        }
    }

    @Test
    public void regionUtils_copyMergedRegionsInRow_byIndexRange() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.createRow(0);
            sheet.createRow(2);
            sheet.createRow(3);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));   // merged region anchored on row 0

            PxlRegionUtils.copyMergedRegionsInRow(sheet, 0, 2, 3);   // (sheet, srcRowIndex, dstStart, dstEnd)

            assertThat(PxlRegionUtils.getMergedRegion(sheet, 2, 0)).isNotNull();
            assertThat(PxlRegionUtils.getMergedRegion(sheet, 3, 0)).isNotNull();
        }
    }

    @Test
    public void sheetUtils_cloneSheet_copiesPrintArea() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Src");
            workbook.setPrintArea(0, "$A$1:$B$2");   // getPrintArea now contains "!" -> the print-area copy branch runs

            final Sheet cloned = PxlSheetUtils.cloneSheet(workbook, 0, "Cloned");

            assertThat(workbook.getSheetName(workbook.getSheetIndex(cloned))).isEqualTo("Cloned");
            assertThat(workbook.getPrintArea(workbook.getSheetIndex(cloned))).isNotNull();
        }
    }

    // ==================================================================
    // PxlFileFormat (root public type)
    // ==================================================================

    @Test
    public void fileFormat_fromPoiWorkbook_poiWorkbookTypes_resolveMatchingFormat() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            assertThat(PxlFileFormat.fromPoiWorkbook(workbook)).isEqualTo(PxlFileFormat.HSSF);
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            assertThat(PxlFileFormat.fromPoiWorkbook(workbook)).isEqualTo(PxlFileFormat.XSSF);
        }

        final SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        try {
            assertThat(PxlFileFormat.fromPoiWorkbook(sxssfWorkbook)).isEqualTo(PxlFileFormat.SXSSF);
        } finally {
            PxlWorkbookUtils.closeWorkbook(sxssfWorkbook);   // disposes the temp files backing the workbook
        }
    }

    @Test
    public void fileFormat_fromPoiWorkbook_streamingWorkbook_resolvesToXssf() throws Exception {
        // The streaming reader opens the same OOXML container, so it reports XSSF - SXSSF is the streaming export format.
        final byte[] xlsx;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("x");
            workbook.write(outputStream);
            xlsx = outputStream.toByteArray();
        }

        try (Workbook streaming = StreamingReader.builder().open(new ByteArrayInputStream(xlsx))) {
            assertThat(PxlFileFormat.fromPoiWorkbook(streaming)).isEqualTo(PxlFileFormat.XSSF);
        }
    }

    @Test
    public void fileFormat_fromPoiWorkbook_nullOrUnknownType_returnsDefaultFormat() throws Exception {
        // A plain lookup: nothing is thrown, an unmatched argument falls back to the default export format.
        assertThat(PxlFileFormat.fromPoiWorkbook(null)).isEqualTo(PxlConstants.DEFAULT_EXPORT_FILE_FORMAT);

        // A Workbook implementation PXL does not know falls back the same way.
        final Workbook unknown = (Workbook) Proxy.newProxyInstance(
                Workbook.class.getClassLoader(), new Class<?>[]{Workbook.class}, (proxy, method, args) -> null);

        assertThat(PxlFileFormat.fromPoiWorkbook(unknown)).isEqualTo(PxlConstants.DEFAULT_EXPORT_FILE_FORMAT);
    }

    @Test
    public void fileFormat_fromWorkbookObject_annotationAbsent_returnsDefaultFormat() throws Exception {
        // A class without @PxlWorkbook carries no declared format, so the default export format stands in.
        assertThat(PxlFileFormat.fromWorkbookObject(Employee.class)).isEqualTo(PxlConstants.DEFAULT_EXPORT_FILE_FORMAT);
    }
}
