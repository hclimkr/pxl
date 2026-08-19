package io.github.hclimkr.pxl.util;

import com.github.pjfanning.xlsx.impl.StreamingSheet;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlWorkbookSupport;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Sheet-level POI helpers: cloning a sheet under a new name, and setting a sheet's print area from row/column
 * indexes or an A1-style range.
 * <p>
 * Cloning goes beyond POI's {@link Workbook#cloneSheet(int)}, which drops page setup: the copy also carries the
 * source sheet's print setup, its fit-to-page and repeating row/column settings, and its print area. The requested
 * name is sanitized with {@link WorkbookUtil#createSafeSheetName(String)} and then made unique within the workbook
 * first, so neither an over-long name, nor one holding characters Excel forbids, nor one another sheet already
 * holds can fail the clone.
 * <p>
 * Both concerns are page-layout metadata a streaming sheet ({@link StreamingSheet}) does not carry, so setting a
 * print area no-ops there, as it does on a {@code null} sheet.
 */
public final class PxlSheetUtils {

    /**
     * The character Excel wraps a sheet name in whenever the name holds something that would otherwise break the
     * reference around it - a space, a comma, an exclamation mark. A quote inside such a name is doubled.
     */
    private static final char SHEET_NAME_QUOTE = '\'';

    /**
     * The character that separates a reference's sheet name from its cell range.
     */
    private static final char SHEET_NAME_SEPARATOR = '!';

    /**
     * The character that separates the ranges of a non-contiguous print area.
     */
    private static final char PRINT_AREA_RANGE_SEPARATOR = ',';

    /**
     * Prevents instantiation.
     */
    private PxlSheetUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Clones the sheet at the given index and gives the copy the supplied name. Beyond POI's
     * {@link Workbook#cloneSheet(int)}, this also copies the source sheet's print setup (paper size,
     * orientation, fit width/height, scale, header/footer margins), its fit-to-page and repeating
     * row/column settings, and its print area. The requested name is sanitized with
     * {@link WorkbookUtil#createSafeSheetName(String)} and then made unique within the workbook before being
     * applied, the way the names of the sheets an export creates are: a name another sheet already holds (POI
     * compares names ignoring case) takes a " (2)", " (3)" ... suffix instead of failing the clone. The name the
     * clone ends up with is therefore not always the one asked for - read it back with
     * {@code workbook.getSheetName(workbook.getSheetIndex(clone))} where it matters.
     * <p>
     * A print area of several ranges is carried over whole, and each range is re-pointed at the clone: POI reports
     * a print area with the source sheet's name in front of every range, while {@code setPrintArea} wants the
     * ranges bare and prefixes them with the destination sheet's name itself.
     *
     * @param workbook        the workbook that owns the source sheet and will hold the clone
     * @param srcSheetIndex   the zero-based index of the sheet to clone
     * @param clonedSheetName the desired name for the cloned sheet
     * @return the cloned sheet
     * @throws PxlNullPointerException if {@code workbook} is {@code null}
     */
    public static Sheet cloneSheet(final Workbook workbook,
                                   final int srcSheetIndex,
                                   final String clonedSheetName)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbook, "workbook");

        final Sheet srcSheet = workbook.getSheetAt(srcSheetIndex);
        final PrintSetup srcPrintSetup = srcSheet.getPrintSetup();

        // The name is settled before the clone joins the workbook: POI hands the clone an interim name of its own
        // ("Src (2)"), and reading that as a collision would push the requested name one number further along.
        final String uniqueSheetName = PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, clonedSheetName);

        final Sheet clonedSheet = workbook.cloneSheet(srcSheetIndex);
        final int clonedSheetIndex = workbook.getSheetIndex(clonedSheet);
        final PrintSetup clonedPrintSetup = clonedSheet.getPrintSetup();

        clonedPrintSetup.setPaperSize(srcPrintSetup.getPaperSize());
        clonedPrintSetup.setLandscape(srcPrintSetup.getLandscape());
        clonedPrintSetup.setFitWidth(srcPrintSetup.getFitWidth());
        clonedPrintSetup.setFitHeight(srcPrintSetup.getFitHeight());
        clonedPrintSetup.setScale(srcPrintSetup.getScale());
        clonedPrintSetup.setHeaderMargin(srcPrintSetup.getHeaderMargin());
        clonedPrintSetup.setFooterMargin(srcPrintSetup.getFooterMargin());

        clonedSheet.setFitToPage(srcSheet.getFitToPage());
        clonedSheet.setRepeatingRows(srcSheet.getRepeatingRows());
        clonedSheet.setRepeatingColumns(srcSheet.getRepeatingColumns());

        workbook.setSheetName(clonedSheetIndex, uniqueSheetName);

        // The rename comes first because setPrintArea stamps the sheet's name, as it stands at that moment, into
        // the reference it stores. Renaming afterwards would leave POI to rewrite that reference on its own.
        final String printArea = workbook.getPrintArea(srcSheetIndex);
        if (StringUtils.isNotBlank(printArea)) {
            workbook.setPrintArea(clonedSheetIndex, removeSheetNamesFromPrintArea(printArea));
        }

        return clonedSheet;
    }

    /**
     * Strips the sheet name from every range of a print area, leaving the bare ranges
     * {@link Workbook#setPrintArea(int, String)} expects.
     * <p>
     * The ranges are split on the separators that sit outside a quoted sheet name, so a name holding a comma
     * keeps its range in one piece.
     * <p>
     * POI splits and renders references itself, but neither half is usable here. {@code AreaReference}'s splitter
     * decides where a quoted name ends by counting the quotes in the segment so far and calling it closed at two,
     * which a doubled quote - the escape for a name such as {@code O'Brien} - pushes past; two such ranges are
     * then read as one and rejected. And rebuilding the survivors through {@code CellReference} would go through
     * {@code AreaReference}, which stores a whole-column reference as rows 1 to 65536 whatever the format, turning
     * an XLSX one into a range 16 times too short. Passing the text through untouched avoids both.
     *
     * @param printArea the print area as {@link Workbook#getPrintArea(int)} reports it
     * @return the same ranges, in the same order, without their sheet names
     */
    private static String removeSheetNamesFromPrintArea(final String printArea) {

        final List<String> ranges = new ArrayList<>();
        boolean quoted = false;
        int rangeStart = 0;

        for (int i = 0; i < printArea.length(); i++) {
            final char character = printArea.charAt(i);

            if (character == SHEET_NAME_QUOTE) {
                quoted = !quoted;
            } else if (!quoted && character == PRINT_AREA_RANGE_SEPARATOR) {
                ranges.add(removeSheetNameFromRange(printArea.substring(rangeStart, i)));
                rangeStart = i + 1;
            }
        }
        ranges.add(removeSheetNameFromRange(printArea.substring(rangeStart)));

        return StringUtils.join(ranges, PRINT_AREA_RANGE_SEPARATOR);
    }

    /**
     * Strips the sheet name from a single range of a print area. The name ends at the first separator that sits
     * outside a quoted name, so a name holding the separator itself is not cut through the middle. A range that
     * names no sheet is returned as it came.
     *
     * @param range one range of a print area, sheet name included
     * @return the range without its sheet name
     */
    private static String removeSheetNameFromRange(final String range) {

        boolean quoted = false;

        for (int i = 0; i < range.length(); i++) {
            final char character = range.charAt(i);

            if (character == SHEET_NAME_QUOTE) {
                quoted = !quoted;
            } else if (!quoted && character == SHEET_NAME_SEPARATOR) {
                return range.substring(i + 1);
            }
        }

        return range;
    }

    /**
     * Sets the sheet's print area from row/column index bounds. A {@code null} sheet or a streaming
     * sheet ({@link StreamingSheet}) is a no-op.
     *
     * @param sheet            the sheet to set the print area on
     * @param startRowIndex    the zero-based first row of the print area
     * @param startColumnIndex the zero-based first column of the print area
     * @param endRowIndex      the zero-based last row of the print area
     * @param endColumnIndex   the zero-based last column of the print area
     */
    public static void setPrintArea(final Sheet sheet,
                                    final int startRowIndex,
                                    final int startColumnIndex,
                                    final int endRowIndex,
                                    final int endColumnIndex) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        final Workbook workbook = sheet.getWorkbook();
        final int sheetIndex = workbook.getSheetIndex(sheet);
        workbook.setPrintArea(sheetIndex, startColumnIndex, endColumnIndex, startRowIndex, endRowIndex);
    }

    /**
     * Sets the sheet's print area from an A1-style range string (for example {@code "A1:D10"}).
     * A {@code null} sheet or a streaming sheet ({@link StreamingSheet}) is a no-op.
     *
     * @param sheet               the sheet to set the print area on
     * @param cellRangeAddressStr the print area as an A1-style range reference
     */
    public static void setPrintArea(final Sheet sheet,
                                    final String cellRangeAddressStr) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        final Workbook workbook = sheet.getWorkbook();
        final int sheetIndex = workbook.getSheetIndex(sheet);
        workbook.setPrintArea(sheetIndex, cellRangeAddressStr);
    }

}
