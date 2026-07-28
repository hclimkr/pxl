package io.github.hclimkr.pxl.util;

import com.github.pjfanning.xlsx.impl.StreamingSheet;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.util.Objects;

/**
 * Sheet-level POI helpers: cloning a sheet under a new name, and setting a sheet's print area from row/column
 * indexes or an A1-style range.
 * <p>
 * Cloning goes beyond POI's {@link Workbook#cloneSheet(int)}, which drops page setup: the copy also carries the
 * source sheet's print setup, its fit-to-page and repeating row/column settings, and its print area. The requested
 * name is sanitized with {@link WorkbookUtil#createSafeSheetName(String)} first, so an over-long name or one
 * holding characters Excel forbids cannot fail the clone.
 * <p>
 * Both concerns are page-layout metadata a streaming sheet ({@link StreamingSheet}) does not carry, so setting a
 * print area no-ops there, as it does on a {@code null} sheet.
 */
public final class PxlSheetUtils {

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
     * {@link WorkbookUtil#createSafeSheetName(String)} before being applied.
     *
     * @param workbook        the workbook that owns the source sheet and will hold the clone
     * @param srcSheetIndex   the zero-based index of the sheet to clone
     * @param clonedSheetName the desired name for the cloned sheet
     * @return the cloned sheet
     */
    public static Sheet cloneSheet(final Workbook workbook,
                                   final int srcSheetIndex,
                                   final String clonedSheetName) {

        final Sheet srcSheet = workbook.getSheetAt(srcSheetIndex);
        final PrintSetup srcPrintSetup = srcSheet.getPrintSetup();

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

        final String printArea = workbook.getPrintArea(srcSheetIndex);
        if (Objects.nonNull(printArea) && printArea.contains("!")) {
            workbook.setPrintArea(clonedSheetIndex, printArea.substring(printArea.indexOf("!") + 1));
        }
        workbook.setSheetName(clonedSheetIndex, WorkbookUtil.createSafeSheetName(clonedSheetName));

        return clonedSheet;
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
