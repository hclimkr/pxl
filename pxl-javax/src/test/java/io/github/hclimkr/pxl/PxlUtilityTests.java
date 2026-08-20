package io.github.hclimkr.pxl;

import com.github.pjfanning.xlsx.StreamingReader;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.styler.PxlVoidStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataCommaSeparatedNumericStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataTextStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataVerticalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderVerticalCenterTextStyler;
import io.github.hclimkr.pxl.tcdata.Employee;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import io.github.hclimkr.pxl.util.*;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Proxy;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.Normalizer;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.regex.Pattern;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Utility/library behavior tests.
 * <p>
 * The first part are reference tests that "do not call the PXL API", moved over from the removed deprecated suite.
 * (regex masking patterns, {@link DecimalFormat}, POI NumberToTextConverter, Unicode normalization, Apache Commons CSV parsing)
 * Ones that originally only printed output were turned into real tests with assertions, and the data is written mostly in English (ASCII).
 * <p>
 * The second part directly exercises the public {@code util/} helpers (PxlMiscUtils, PxlRegionUtils, PxlSheetUtils,
 * PxlRowUtils, PxlCellUtils, PxlCollectionUtils) against a plain XSSF workbook, since those are consumer-facing
 * entry points, and closes with the public stylers and the {@code PxlFileFormat} / {@code PxlExcelEngine} lookups.
 * <p>
 * Tests that reach into {@code internal/} live in {@link PxlInternalTests} instead - everything here is public API.
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
    // POI NumberToTextConverter (renders without scientific notation - the basis for PXL's number->string rendering)
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
    public void miscUtils_columnIndexAndLetters_convertBothWays() throws PxlException {
        assertThat(PxlMiscUtils.convertColumnIndexToColumnString(0)).isEqualTo("A");
        assertThat(PxlMiscUtils.convertColumnIndexToColumnString(26)).isEqualTo("AA");
        assertThat(PxlMiscUtils.convertColumnStringToColumnIndex("A")).isEqualTo(0);
        assertThat(PxlMiscUtils.convertColumnStringToColumnIndex("AA")).isEqualTo(26);
        // An absolute-reference marker is part of the letters POI accepts.
        assertThat(PxlMiscUtils.convertColumnStringToColumnIndex("$B")).isEqualTo(1);
    }

    @Test
    public void miscUtils_columnStringToIndex_missingOrMalformed_throws() {
        // A required argument is checked before POI sees it.
        assertThrows(PxlNullPointerException.class, () -> PxlMiscUtils.convertColumnStringToColumnIndex(null));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertColumnStringToColumnIndex(" "));
        // POI answers a nonsense index for anything that is not column letters ("1" -> -16), so the shape is
        // rejected up front instead.
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertColumnStringToColumnIndex("1"));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertColumnStringToColumnIndex("A1"));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertColumnStringToColumnIndex("A$B"));
    }

    @Test
    public void miscUtils_cellReferenceAndRange_formatted() throws PxlException {
        assertThat(PxlMiscUtils.convertIndexesToCellReferenceString(2, 1)).isEqualTo("B3");
        assertThat(PxlMiscUtils.convertIndexesToCellRangeAddressString(0, 0, 9, 3)).isEqualTo("A1:D10");
    }

    @Test
    public void miscUtils_negativeIndexes_rejectedBeforePoiAnswersWithAValue() {
        // POI reads -1 as "not stated" and builds half a reference from it - (-1, 0) as "A", (0, -1) as "1" -
        // and answers a negative column index with an empty string. A wrong reference is harder to notice than
        // a refusal, so the shape is settled before POI sees it.
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertColumnIndexToColumnString(-1));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertIndexesToCellReferenceString(-1, 0));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertIndexesToCellReferenceString(0, -1));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertIndexesToCellRangeAddressString(-1, 0, 9, 3));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertIndexesToCellRangeAddressString(0, -1, 9, 3));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertIndexesToCellRangeAddressString(0, 0, -1, 3));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertIndexesToCellRangeAddressString(0, 0, 9, -1));
    }

    @Test
    public void miscUtils_cellReferenceToIndexes_parsedAndValidated() throws PxlException {
        final Pair<Integer, Integer> indexes = PxlMiscUtils.convertCellReferenceStringToIndexes("B3");
        assertThat(indexes.getLeft()).isEqualTo(2);
        assertThat(indexes.getRight()).isEqualTo(1);
        // A column-only reference has no row -> rejected.
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertCellReferenceStringToIndexes("B"));
        // A required argument is checked before POI sees it.
        assertThrows(PxlNullPointerException.class, () -> PxlMiscUtils.convertCellReferenceStringToIndexes(null));
        assertThrows(PxlArgumentException.class, () -> PxlMiscUtils.convertCellReferenceStringToIndexes(" "));
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
    public void regionUtils_copyMergedRegionsInRow_sameRow_isNoOp() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Row srcRow = sheet.createRow(5);
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));   // A6:B6, anchored on row 5

            // Re-creating a region where it already sits would overlap the original, so both overloads no-op
            // instead of letting POI throw.
            PxlRegionUtils.copyMergedRegionsInRow(sheet, srcRow, srcRow);
            PxlRegionUtils.copyMergedRegionsInRow(sheet, 5, 5);

            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactly("A6:B6");
        }
    }

    @Test
    public void regionUtils_copyMergedRegionsInRow_rangeCoveringSourceRow_skipsThatRowOnly() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            for (int rowIndex = 4; rowIndex <= 7; rowIndex++) {
                sheet.createRow(rowIndex);
            }
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));   // A6:B6, anchored on row 5

            PxlRegionUtils.copyMergedRegionsInRow(sheet, 5, 4, 7);   // the range covers the source row

            // Row 5 is skipped; the rest of the range is copied rather than the whole call failing on it.
            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactlyInAnyOrder("A6:B6", "A5:B5", "A7:B7", "A8:B8");
        }
    }

    @Test
    public void regionUtils_copyMergedRegionsInRow_rangeWithMultiRowRegion_skipsOnlyThatRegion() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.createRow(5);
            for (int rowIndex = 10; rowIndex <= 12; rowIndex++) {
                sheet.createRow(rowIndex);
            }
            sheet.addMergedRegion(new CellRangeAddress(5, 7, 0, 0));   // A6:A8, three rows tall
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 2));   // B6:C6, one row tall

            PxlRegionUtils.copyMergedRegionsInRow(sheet, 5, 10, 12);

            // Destinations are one row apart, so replicating the three-row region would overlap the copy made
            // for the row before it; it is skipped, while the one-row region is copied onto every row.
            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactlyInAnyOrder("A6:A8", "B6:C6", "B11:C11", "B12:C12", "B13:C13");
        }
    }

    @Test
    public void regionUtils_copyMergedRegionsInRow_singleDestination_replicatesMultiRowRegion() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Row srcRow = sheet.createRow(5);
            final Row dstRow = sheet.createRow(20);
            sheet.addMergedRegion(new CellRangeAddress(5, 7, 0, 0));   // A6:A8, three rows tall

            // A single destination cannot collide with a copy of itself, so the row span is carried over.
            PxlRegionUtils.copyMergedRegionsInRow(sheet, srcRow, dstRow);

            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactlyInAnyOrder("A6:A8", "A21:A23");
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
    public void rowUtils_getRow_negativeIndex_nullWhenReadingAndThrowsWhenCreating() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");

            // Reading answers a bad index the way it answers an absent row.
            assertThat(PxlRowUtils.getRow(sheet, -1, false)).isNull();

            // Creating a row at one is a change to the sheet, so it is refused rather than left to POI - and the
            // refusal does not depend on the sheet being there.
            assertThrows(PxlArgumentException.class, () -> PxlRowUtils.getRow(sheet, -1, true));
            assertThrows(PxlArgumentException.class, () -> PxlRowUtils.getRow(null, -1, true));
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
    public void rowUtils_copyRow_sameRowIndex_isNoOp() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            for (int rowIndex = 0; rowIndex <= 2; rowIndex++) {
                sheet.createRow(rowIndex).createCell(0).setCellValue("r" + rowIndex);
            }
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 1));

            PxlRowUtils.copyRow(sheet, 1, 1);   // copying a row onto itself has nothing to do

            assertThat(sheet.getLastRowNum()).isEqualTo(2);
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("r1");
            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("r2");
            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactly("A2:B2");
        }
    }

    @Test
    public void rowUtils_copyRow_rowWithoutRowStyle_copiesOnXls() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            for (int rowIndex = 0; rowIndex <= 2; rowIndex++) {
                sheet.createRow(rowIndex).createCell(0).setCellValue("r" + rowIndex);
            }

            // HSSFRow.setRowStyle(null) throws, unlike its XSSF counterpart, so a row that never got a style
            // must not be handed one.
            assertThat(sheet.getRow(0).getRowStyle()).isNull();

            PxlRowUtils.copyRow(sheet, 0, 2);

            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("r0");
        }
    }

    @Test
    public void rowUtils_copyRow_xlsDestinationBeforeSource_copiesValueAndRegion() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            for (int rowIndex = 0; rowIndex <= 6; rowIndex++) {
                sheet.createRow(rowIndex).createCell(0).setCellValue("r" + rowIndex);
                sheet.getRow(rowIndex).createCell(1).setCellValue("x" + rowIndex);
            }
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));   // A6:B6, anchored on the source row

            PxlRowUtils.copyRow(sheet, 5, 3);   // the shift pushes the source down to row 6

            // HSSF's shiftRows empties the source Row object instead of moving it, so a reference taken before
            // the shift would have no cells left and the destination would come out blank.
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("r5");
            assertThat(sheet.getRow(6).getCell(0).getStringCellValue()).isEqualTo("r5");   // the source itself
            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactlyInAnyOrder("A7:B7", "A4:B4");
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
    public void cellUtils_getCell_negativeIndex_nullWhenReadingAndThrowsWhenCreating() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("v");

            // A read yields null for either index - POI answers a negative column with a raw
            // IllegalArgumentException even when it is only being read.
            assertThat(PxlCellUtils.getCell(sheet, -1, 0, false)).isNull();
            assertThat(PxlCellUtils.getCell(sheet, 0, -1, false)).isNull();
            assertThat(PxlCellUtils.getCellWithMerges(sheet, -1, -1)).isNull();

            // Creating one is refused, and the row index is reported first.
            assertThat(assertThrows(PxlArgumentException.class, () -> PxlCellUtils.getCell(sheet, -1, -1, true)))
                    .hasMessageContaining("rowIndex");
            assertThat(assertThrows(PxlArgumentException.class, () -> PxlCellUtils.getCell(sheet, 0, -1, true)))
                    .hasMessageContaining("columnIndex");

            // The whole setCellValue family reaches the sheet through that one lookup.
            assertThrows(PxlArgumentException.class, () -> PxlCellUtils.setCellValue(sheet, 0, -1, "x", true));
            assertThrows(PxlArgumentException.class, () -> PxlCellUtils.setCellFormula(sheet, -1, 0, "1+2", true));
            assertThrows(PxlArgumentException.class, () -> PxlCellUtils.setCellBlank(sheet, -1, 0, true));
            assertThrows(PxlArgumentException.class, () -> PxlCellUtils.copyCell(sheet, 0, 0, -1, 0, true));
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

            // createIfNone governs the destination: without it an absent destination is left alone.
            assertThat(PxlCellUtils.copyCell(sheet, 0, 0, 4, 4, false)).isNull();
            assertThat(sheet.getRow(4)).isNull();

            // An existing destination is copied onto even when nothing may be created.
            sheet.createRow(5).createCell(5).setCellValue("overwrite me");
            final Cell dstCell3 = PxlCellUtils.copyCell(sheet, 0, 0, 5, 5, false);
            assertThat(dstCell3).isNotNull();
            assertThat(dstCell3.getStringCellValue()).isEqualTo("copyme");
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
    public void cellUtils_cloneCellStyle_nullArgument_returnsNull() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Cell cell = workbook.createSheet("S").createRow(0).createCell(0);

            // A missing cell or target workbook is answered the way a style the workbook cannot take is: with
            // null, rather than a raw NullPointerException escaping a public utility.
            assertThat(PxlCellUtils.cloneCellStyle(null)).isNull();
            assertThat(PxlCellUtils.cloneCellStyle(null, workbook)).isNull();
            assertThat(PxlCellUtils.cloneCellStyle(cell, null)).isNull();
        }
    }

    @Test
    public void cellUtils_cloneCellStyle_styleLimitExhausted_returnsNullWithoutThrowing() throws Exception {
        try (HSSFWorkbook srcWorkbook = new HSSFWorkbook();
             HSSFWorkbook dstWorkbook = new HSSFWorkbook()) {

            final Cell srcCell = srcWorkbook.createSheet("S").createRow(0).createCell(0);

            // Use up the destination workbook's cell-style pool so POI refuses to hand out another style.
            assertThrows(IllegalStateException.class, () -> {
                for (int i = 0; i < 5000; i++) {
                    dstWorkbook.createCellStyle();
                }
            });

            // The refusal is reported through an SLF4J WARN instead of escaping, and nothing half-built comes back.
            assertThat(PxlCellUtils.cloneCellStyle(srcCell, dstWorkbook)).isNull();
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

    @Test
    public void cellUtils_addNoteToCell_xls_setsComment() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Cell cell = workbook.createSheet("S").createRow(0).createCell(0);

            // The note inset is given in pixels; XLS states an anchor offset as a fraction of the anchored cell
            // instead, and POI rejects anything outside that fraction's range. Passing the pixel count through as
            // if it were EMU used to fail here with "dx1 must be between 0 and 1023".
            PxlCellUtils.addNoteToCell(cell, "hello");

            assertThat(cell.getCellComment()).isNotNull();
            assertThat(cell.getCellComment().getAuthor()).isEqualTo(PxlConstants.PXL_CREATOR);

            final ClientAnchor anchor = cell.getCellComment().getClientAnchor();
            assertThat(anchor.getDx1()).isBetween(0, 1023);
            assertThat(anchor.getDx2()).isBetween(0, 1023);
            assertThat(anchor.getDy1()).isBetween(0, 255);
            assertThat(anchor.getDy2()).isBetween(0, 255);
        }
    }

    @Test
    public void cellUtils_addHyperlinkToCell_resolvesTypeFromAddress() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Row row = workbook.createSheet("S").createRow(0);

            // A null cell and a blank address are no-ops.
            PxlCellUtils.addHyperlinkToCell(null, "https://example.com/");
            final Cell blankCell = row.createCell(0);
            PxlCellUtils.addHyperlinkToCell(blankCell, "  ");
            assertThat(blankCell.getHyperlink()).isNull();

            // The type is read off the address when none is named.
            assertThat(resolvedHyperlink(row, 1, "https://example.com/a").getType()).isEqualTo(HyperlinkType.URL);
            assertThat(resolvedHyperlink(row, 2, "MailTo:someone@example.com").getType()).isEqualTo(HyperlinkType.EMAIL);
            assertThat(resolvedHyperlink(row, 3, "#'S'!A1").getType()).isEqualTo(HyperlinkType.DOCUMENT);
            assertThat(resolvedHyperlink(row, 4, "docs/readme.txt").getType()).isEqualTo(HyperlinkType.FILE);

            // A file:// URL carries a scheme, so it reads as a URL rather than a file path.
            assertThat(resolvedHyperlink(row, 5, "file://server/share/a.txt").getType()).isEqualTo(HyperlinkType.URL);

            // Naming a type skips the reading.
            final Cell namedCell = row.createCell(6);
            PxlCellUtils.addHyperlinkToCell(namedCell, "docs/readme.txt", null, HyperlinkType.URL);
            assertThat(namedCell.getHyperlink().getType()).isEqualTo(HyperlinkType.URL);

            // The link is laid over the cell: the value it already held stays, and the label is set only when given.
            final Cell valuedCell = row.createCell(7);
            valuedCell.setCellValue("Example");
            PxlCellUtils.addHyperlinkToCell(valuedCell, "https://example.com/", "Example site");
            assertThat(valuedCell.getStringCellValue()).isEqualTo("Example");
            assertThat(valuedCell.getHyperlink().getAddress()).isEqualTo("https://example.com/");
            assertThat(valuedCell.getHyperlink().getLabel()).isEqualTo("Example site");
            assertThat(resolvedHyperlink(row, 8, "https://example.com/").getLabel()).isNull();
        }
    }

    @Test
    public void cellUtils_addHyperlinkToCell_documentAddress_dropsLeadingHash() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Row row = workbook.createSheet("S").createRow(0);

            // POI takes the location alone, so the '#' that marks the address as internal is dropped.
            assertThat(resolvedHyperlink(row, 0, "#'S'!A1").getAddress()).isEqualTo("'S'!A1");

            // A named DOCUMENT type drops it as well, and an address that is nothing but the '#' is a no-op.
            final Cell namedCell = row.createCell(1);
            PxlCellUtils.addHyperlinkToCell(namedCell, "#DefinedName", null, HyperlinkType.DOCUMENT);
            assertThat(namedCell.getHyperlink().getAddress()).isEqualTo("DefinedName");

            final Cell hashOnlyCell = row.createCell(2);
            PxlCellUtils.addHyperlinkToCell(hashOnlyCell, "#");
            assertThat(hashOnlyCell.getHyperlink()).isNull();
        }
    }

    @Test
    public void cellUtils_addHyperlinkToCell_typeNone_throws() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Cell cell = workbook.createSheet("S").createRow(0).createCell(0);

            // NONE is not a link a cell can carry: it is refused before POI raises its own exception, and it is
            // refused for a null cell too, since the argument is checked first.
            assertThrows(PxlArgumentException.class,
                    () -> PxlCellUtils.addHyperlinkToCell(cell, "https://example.com/", null, HyperlinkType.NONE));
            assertThrows(PxlArgumentException.class,
                    () -> PxlCellUtils.addHyperlinkToCell(null, "https://example.com/", null, HyperlinkType.NONE));
            assertThat(cell.getHyperlink()).isNull();
        }
    }

    @Test
    public void cellUtils_addHyperlinkToCell_xls_setsHyperlink() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Cell cell = workbook.createSheet("S").createRow(0).createCell(0);
            cell.setCellValue("Example");

            PxlCellUtils.addHyperlinkToCell(cell, "https://example.com/", "Example site");

            assertThat(cell.getHyperlink()).isNotNull();
            assertThat(cell.getHyperlink().getType()).isEqualTo(HyperlinkType.URL);
            assertThat(cell.getHyperlink().getAddress()).isEqualTo("https://example.com/");

            // XLS keeps a fixed moniker per link type, so the label given here does not reach the file (POI 5.5.1).
            assertThat(cell.getHyperlink().getLabel()).isEqualTo("url");
        }
    }

    /**
     * Links a fresh cell in the row through PxlCellUtils, letting the address pick the type, and hands the
     * attached hyperlink back.
     */
    private static Hyperlink resolvedHyperlink(final Row row,
                                               final int columnIndex,
                                               final String address) {

        final Cell cell = row.createCell(columnIndex);
        PxlCellUtils.addHyperlinkToCell(cell, address);

        return cell.getHyperlink();
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

    /**
     * Attaches a comment spanning 2 columns and 3 rows, so a copy's anchor size can be checked against the source.
     */
    private static Comment attachComment(final Cell cell,
                                         final String text,
                                         final String author) {

        final Sheet sheet = cell.getSheet();
        final CreationHelper creationHelper = sheet.getWorkbook().getCreationHelper();

        final ClientAnchor anchor = creationHelper.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + 2);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + 3);

        final Comment comment = sheet.createDrawingPatriarch().createCellComment(anchor);
        comment.setString(creationHelper.createRichTextString(text));
        comment.setAuthor(author);
        cell.setCellComment(comment);

        return comment;
    }

    private static Hyperlink attachHyperlink(final Cell cell,
                                             final String address,
                                             final String label) {

        final Hyperlink hyperlink = cell.getSheet()
                .getWorkbook()
                .getCreationHelper()
                .createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(address);
        hyperlink.setLabel(label);
        cell.setHyperlink(hyperlink);

        return hyperlink;
    }

    @Test
    public void cellUtils_copyCell_sameWorkbook_keepsCommentAndHyperlinkOnEveryCell() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Cell srcCell = sheet.createRow(0).createCell(0);
            srcCell.setCellValue("template");
            final Comment srcComment = attachComment(srcCell, "note", "author-A");
            final Hyperlink srcHyperlink = attachHyperlink(srcCell, "https://example.com/x", "label");

            final Cell dstCell1 = sheet.createRow(1).createCell(0);
            final Cell dstCell2 = sheet.createRow(2).createCell(0);
            PxlCellUtils.copyCell(srcCell, dstCell1);
            PxlCellUtils.copyCell(srcCell, dstCell2);

            // Handing POI the source's own objects would relocate them, emptying the source and every destination but one.
            assertThat(srcCell.getCellComment()).isNotNull();
            assertThat(srcCell.getHyperlink()).isNotNull();
            assertThat(sheet.getCellComments().keySet()).containsExactlyInAnyOrder(
                    new CellAddress(0, 0), new CellAddress(1, 0), new CellAddress(2, 0));

            for (final Cell dstCell : Arrays.asList(dstCell1, dstCell2)) {
                assertThat(dstCell.getCellComment()).isNotSameAs(srcComment);
                assertThat(dstCell.getCellComment().getString().getString()).isEqualTo("note");
                assertThat(dstCell.getCellComment().getAuthor()).isEqualTo("author-A");
                assertThat(dstCell.getHyperlink()).isNotSameAs(srcHyperlink);
                assertThat(dstCell.getHyperlink().getAddress()).isEqualTo("https://example.com/x");
                assertThat(dstCell.getHyperlink().getLabel()).isEqualTo("label");
            }

            // The copied anchor keeps the source's size while sitting over the destination cell.
            final ClientAnchor dstAnchor = dstCell1.getCellComment().getClientAnchor();
            assertThat(dstAnchor.getCol1()).isZero();
            assertThat(dstAnchor.getCol2() - dstAnchor.getCol1()).isEqualTo(2);
            assertThat(dstAnchor.getRow1()).isEqualTo(1);
            assertThat(dstAnchor.getRow2() - dstAnchor.getRow1()).isEqualTo(3);
        }
    }

    @Test
    public void cellUtils_copyCell_crossWorkbook_commentAndHyperlinkBelongToDestination() throws Exception {
        try (XSSFWorkbook srcWorkbook = new XSSFWorkbook(); XSSFWorkbook dstWorkbook = new XSSFWorkbook()) {
            final Sheet srcSheet = srcWorkbook.createSheet("S");
            final Cell srcCell = srcSheet.createRow(0).createCell(0);
            srcCell.setCellValue("template");
            attachComment(srcCell, "note", "author-A");
            attachHyperlink(srcCell, "https://example.com/x", "label");

            final Sheet dstSheet = dstWorkbook.createSheet("T");
            final Cell dstCell = dstSheet.createRow(4).createCell(2);
            PxlCellUtils.copyCell(srcCell, dstCell);

            // The source workbook keeps its comment where it was, rather than having it moved to the destination's address.
            assertThat(srcSheet.getCellComments().keySet()).containsExactly(new CellAddress(0, 0));
            assertThat(dstSheet.getCellComments().keySet()).containsExactly(new CellAddress(4, 2));
            assertThat(dstCell.getCellComment().getString().getString()).isEqualTo("note");

            // The hyperlink is not shared between the workbooks, so editing the copy leaves the source alone.
            dstCell.getHyperlink().setAddress("https://example.com/changed");
            assertThat(srcCell.getHyperlink().getAddress()).isEqualTo("https://example.com/x");
        }
    }

    @Test
    public void cellUtils_copyCell_hssf_anchorsOneCommentPerCell() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            final Cell srcCell = sheet.createRow(0).createCell(0);
            srcCell.setCellValue("template");
            attachComment(srcCell, "note", "author-A");
            attachHyperlink(srcCell, "https://example.com/x", "label");

            PxlCellUtils.copyCell(srcCell, sheet.createRow(1).createCell(0));
            PxlCellUtils.copyCell(srcCell, sheet.createRow(2).createCell(0));

            // An HSSFCell answers with the comment reference it caches, so only the sheet's registry shows where the
            // notes really sit - and that is what a saved file holds.
            assertThat(sheet.getCellComments().keySet()).containsExactlyInAnyOrder(
                    new CellAddress(0, 0), new CellAddress(1, 0), new CellAddress(2, 0));
            assertThat(sheet.getRow(1).getCell(0).getHyperlink().getAddress()).isEqualTo("https://example.com/x");
            assertThat(sheet.getRow(2).getCell(0).getHyperlink().getAddress()).isEqualTo("https://example.com/x");
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
    public void rowUtils_copyRowMultiplyByRange_rangeCoveringSourceRow_fillsEveryRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            for (int rowIndex = 0; rowIndex <= 6; rowIndex++) {
                sheet.createRow(rowIndex).createCell(0).setCellValue("r" + rowIndex);
            }
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));   // A6:B6, anchored on the source row

            PxlRowUtils.copyRowMultiplyByRange(sheet, 5, 3, 5);   // the range covers the source row index

            // The shift moves the source row to index 8 first, so index 5 is an emptied destination that still
            // has to be filled - the guard compares shifted coordinates rather than the original index.
            for (int rowIndex = 3; rowIndex <= 5; rowIndex++) {
                assertThat(sheet.getRow(rowIndex).getCell(0).getStringCellValue()).isEqualTo("r5");
            }
            assertThat(sheet.getRow(8).getCell(0).getStringCellValue()).isEqualTo("r5");   // the source itself
            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactlyInAnyOrder("A9:B9", "A4:B4", "A5:B5", "A6:B6");
        }
    }

    @Test
    public void rowUtils_copyRowMultiplyByRange_xlsRangeCoveringSourceRow_copiesEveryRow() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            for (int rowIndex = 0; rowIndex <= 6; rowIndex++) {
                sheet.createRow(rowIndex).createCell(0).setCellValue("r" + rowIndex);
                sheet.getRow(rowIndex).createCell(1).setCellValue("x" + rowIndex);
            }
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));   // A6:B6, anchored on the source row

            PxlRowUtils.copyRowMultiplyByRange(sheet, 5, 3, 5);   // the shift pushes the source down to row 8

            // Same expectation as the XSSF case above: XLS used to produce three blank rows here, because
            // HSSF's shiftRows moves the cells out of the source Row object rather than moving the object.
            for (int rowIndex = 3; rowIndex <= 5; rowIndex++) {
                assertThat(sheet.getRow(rowIndex).getCell(0).getStringCellValue()).isEqualTo("r5");
            }
            assertThat(sheet.getRow(8).getCell(0).getStringCellValue()).isEqualTo("r5");   // the source itself
            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactlyInAnyOrder("A9:B9", "A4:B4", "A5:B5", "A6:B6");
        }
    }

    @Test
    public void rowUtils_copyRowMultiplyByRange_multiRowMergedRegion_copiesRowsWithoutThatRegion() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            for (int rowIndex = 0; rowIndex <= 7; rowIndex++) {
                sheet.createRow(rowIndex).createCell(0).setCellValue("r" + rowIndex);
            }
            sheet.addMergedRegion(new CellRangeAddress(5, 7, 0, 0));   // A6:A8, anchored on the source row

            // This used to fail: the second destination row got a region overlapping the first one's, and POI
            // rejected it with an IllegalStateException once part of the range had already been merged.
            PxlRowUtils.copyRowMultiplyByRange(sheet, 5, 10, 12);

            for (int rowIndex = 10; rowIndex <= 12; rowIndex++) {
                assertThat(sheet.getRow(rowIndex).getCell(0).getStringCellValue()).isEqualTo("r5");
            }
            assertThat(sheet.getMergedRegions()).extracting(CellRangeAddress::formatAsString)
                    .containsExactly("A6:A8");
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
            workbook.setPrintArea(0, "$A$1:$B$2");   // POI keeps this as "Src!$A$1:$B$2"

            final Sheet cloned = PxlSheetUtils.cloneSheet(workbook, 0, "Cloned");

            assertThat(workbook.getSheetName(workbook.getSheetIndex(cloned))).isEqualTo("Cloned");
            // The range comes over pointing at the clone, not at the sheet it was read from.
            assertThat(workbook.getPrintArea(workbook.getSheetIndex(cloned))).isEqualTo("Cloned!$A$1:$B$2");
        }
    }

    @Test
    public void sheetUtils_cloneSheet_multiRangePrintArea_pointsEveryRangeAtTheClone() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Src");
            workbook.setPrintArea(0, "$A$1:$B$2,$D$1:$E$2");   // POI names the source sheet in front of each range

            final Sheet cloned = PxlSheetUtils.cloneSheet(workbook, 0, "Cloned");

            // Dropping only the first sheet name would leave the second range printing the source sheet's cells.
            assertThat(workbook.getPrintArea(workbook.getSheetIndex(cloned)))
                    .isEqualTo("Cloned!$A$1:$B$2,Cloned!$D$1:$E$2");
        }
    }

    @Test
    public void sheetUtils_cloneSheet_quotedSheetNamePrintArea_survives() throws Exception {
        // Excel forbids none of ',', '!' and a non-leading, non-trailing '\'' in a sheet name, so POI quotes the
        // name instead of rejecting it and doubles the quote inside: the print area reads
        // 'A,B!C''D'!$A$1:$B$2,'A,B!C''D'!$D$1:$E$2. Every separator inside those quotes, the doubled quote
        // included, has to be told apart from the ones doing the actual separating.
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("A,B!C'D");
            workbook.setPrintArea(0, "$A$1:$B$2,$D$1:$E$2");

            final Sheet cloned = PxlSheetUtils.cloneSheet(workbook, 0, "Cloned");

            assertThat(workbook.getPrintArea(workbook.getSheetIndex(cloned)))
                    .isEqualTo("Cloned!$A$1:$B$2,Cloned!$D$1:$E$2");
        }
    }

    @Test
    public void sheetUtils_cloneSheet_nameAlreadyTaken_suffixesInsteadOfFailing() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Src");
            workbook.createSheet("Copy");

            final Sheet cloned = PxlSheetUtils.cloneSheet(workbook, 0, "Copy");

            // POI's setSheetName turns down a name another sheet holds; the clone takes the next free number.
            assertThat(workbook.getSheetName(workbook.getSheetIndex(cloned))).isEqualTo("Copy (2)");
            assertThat(workbook.getSheetName(1)).isEqualTo("Copy");
        }
    }

    @Test
    public void sheetUtils_cloneSheet_nameTakenInAnotherCase_suffixesInsteadOfFailing() throws Exception {
        // POI turns the name down whatever case it is asked in, so the uniqueness check ignores case as well.
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Src");
            workbook.createSheet("Copy");

            final Sheet cloned = PxlSheetUtils.cloneSheet(workbook, 0, "copy");

            assertThat(workbook.getSheetName(workbook.getSheetIndex(cloned))).isEqualTo("copy (2)");
        }
    }

    @Test
    public void sheetUtils_cloneSheet_nameMatchingTheClonesInterimName_keepsIt() throws Exception {
        // POI names the clone "Src (2)" as it joins the workbook. Settling the name only afterwards would read that
        // interim name as a collision and push the requested one on to "Src (2) (2)".
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Src");

            final Sheet cloned = PxlSheetUtils.cloneSheet(workbook, 0, "Src (2)");

            assertThat(workbook.getSheetName(workbook.getSheetIndex(cloned))).isEqualTo("Src (2)");
        }
    }

    // ==================================================================
    // PxlFileFormat (root public type) - the physical format axis
    // ==================================================================

    @Test
    public void fileFormat_fromPoiWorkbook_poiWorkbookTypes_resolveWrittenFormat() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            assertThat(PxlFileFormat.fromPoiWorkbook(workbook)).isEqualTo(PxlFileFormat.XLS);
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            assertThat(PxlFileFormat.fromPoiWorkbook(workbook)).isEqualTo(PxlFileFormat.XLSX);
        }

        // SXSSF is a different writer, not a different format - it produces the same XLSX container as XSSF.
        final SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        try {
            assertThat(PxlFileFormat.fromPoiWorkbook(sxssfWorkbook)).isEqualTo(PxlFileFormat.XLSX);
        } finally {
            PxlWorkbookUtils.closeWorkbook(sxssfWorkbook);   // disposes the temp files backing the workbook
        }
    }

    @Test
    public void fileFormat_fromPoiWorkbook_streamingWorkbook_resolvesToXlsx() throws Exception {
        // The streaming reader opens the same OOXML container, so it holds the XLSX format like the other two.
        final byte[] xlsx;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("x");
            workbook.write(outputStream);
            xlsx = outputStream.toByteArray();
        }

        try (Workbook streaming = StreamingReader.builder().open(new ByteArrayInputStream(xlsx))) {
            assertThat(PxlFileFormat.fromPoiWorkbook(streaming)).isEqualTo(PxlFileFormat.XLSX);
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

    // ==================================================================
    // PxlExcelEngine (root public type) - the writer axis
    // ==================================================================

    @Test
    public void excelEngine_fromPoiWorkbook_poiWriterTypes_resolveMatchingEngine() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            assertThat(PxlExcelEngine.fromPoiWorkbook(workbook)).isEqualTo(PxlExcelEngine.HSSF);
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            assertThat(PxlExcelEngine.fromPoiWorkbook(workbook)).isEqualTo(PxlExcelEngine.XSSF);
        }

        // Here the two XLSX writers are told apart - that is what this axis is for.
        final SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        try {
            assertThat(PxlExcelEngine.fromPoiWorkbook(sxssfWorkbook)).isEqualTo(PxlExcelEngine.SXSSF);
        } finally {
            PxlWorkbookUtils.closeWorkbook(sxssfWorkbook);   // disposes the temp files backing the workbook
        }
    }

    @Test
    public void excelEngine_fromPoiWorkbook_nullOrNonWriter_returnsDefaultEngine() throws Exception {
        // A plain lookup: nothing is thrown, an unmatched argument falls back to the default export engine.
        assertThat(PxlExcelEngine.fromPoiWorkbook(null)).isEqualTo(PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE);

        // A Workbook implementation PXL does not know falls back the same way.
        final Workbook unknown = (Workbook) Proxy.newProxyInstance(
                Workbook.class.getClassLoader(), new Class<?>[]{Workbook.class}, (proxy, method, args) -> null);

        assertThat(PxlExcelEngine.fromPoiWorkbook(unknown)).isEqualTo(PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE);
    }

    @Test
    public void excelEngine_fromWorkbookObject_annotationAbsent_returnsDefaultEngine() throws Exception {
        // A class without @PxlWorkbook declares no engine, so the default export engine stands in.
        assertThat(PxlExcelEngine.fromWorkbookObject(Employee.class)).isEqualTo(PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE);
    }
}
