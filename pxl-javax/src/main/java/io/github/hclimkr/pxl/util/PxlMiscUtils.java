package io.github.hclimkr.pxl.util;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.styler.PxlStyler;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.Objects;

/**
 * Miscellaneous spreadsheet helpers: conversions between numeric row/column indexes and A1-style
 * column letters, cell references and range addresses, plus a check for whether a styler class is
 * an effective (usable) cell styler.
 */
public final class PxlMiscUtils {

    /**
     * Prevents instantiation.
     */
    private PxlMiscUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Converts a zero-based column index to its A1-style column letters (for example {@code 0 -> "A"},
     * {@code 26 -> "AA"}).
     *
     * @param columnIndex the zero-based column index
     * @return the column letters
     */
    public static String convertColumnIndexToColumnString(final int columnIndex) {

        return CellReference.convertNumToColString(columnIndex);
    }

    /**
     * Converts A1-style column letters (for example {@code "A"}, {@code "AA"}) to a zero-based column index.
     *
     * @param columnString the column letters
     * @return the zero-based column index
     */
    public static int convertColumnStringToColumnIndex(final String columnString) {

        return CellReference.convertColStringToIndex(columnString);
    }

    /**
     * Builds an A1-style cell reference (for example {@code "B3"}) from zero-based row/column indexes,
     * formatted without absolute-reference (`$`) markers.
     *
     * @param rowIndex    the zero-based row index
     * @param columnIndex the zero-based column index
     * @return the A1-style cell reference
     */
    public static String convertIndexesToCellReferenceString(final int rowIndex,
                                                             final int columnIndex) {

        final CellReference cellRef = new CellReference(rowIndex, columnIndex);

        return cellRef.formatAsString(false);
    }

    /**
     * Builds an A1-style range address (for example {@code "A1:D10"}) from zero-based start/end
     * row and column indexes.
     *
     * @param startRowIndex    the zero-based first row of the range
     * @param startColumnIndex the zero-based first column of the range
     * @param endRowIndex      the zero-based last row of the range
     * @param endColumnIndex   the zero-based last column of the range
     * @return the A1-style range address
     */
    public static String convertIndexesToCellRangeAddressString(final int startRowIndex,
                                                                int startColumnIndex,
                                                                int endRowIndex,
                                                                int endColumnIndex) {

        final CellRangeAddress cellRangeAddress = new CellRangeAddress(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);

        return cellRangeAddress.formatAsString();
    }

    /**
     * Parses an A1-style cell reference into a (row index, column index) pair, both zero-based. The
     * reference must carry both a row and a column; a column-only or row-only reference is rejected.
     *
     * @param cellRefStr the A1-style cell reference
     * @return a pair of (zero-based row index, zero-based column index)
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Pair<Integer, Integer> convertCellReferenceStringToIndexes(final String cellRefStr)
            throws PxlArgumentException {

        final CellReference cellRef = new CellReference(cellRefStr);

        if (cellRef.getRow() < 0 || cellRef.getCol() < 0) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_CELL_REF_INVALID, cellRefStr));
        }

        return Pair.of(cellRef.getRow(), (int) cellRef.getCol());
    }

    /**
     * Determines whether the given styler class is an effective (concrete, usable) cell styler. It is
     * not effective when {@code null}, when it is an interface, or when it is the sentinel
     * {@link PxlConstants#VOID_CELL_STYLER} placeholder.
     *
     * @param cellStylerClass the styler class to test, may be {@code null}
     * @return {@code true} if the class is a concrete, non-void styler
     */
    public static boolean isEffectiveCellStylerClass(final Class<? extends PxlStyler> cellStylerClass) {

        if (Objects.isNull(cellStylerClass)) {
            return false;
        }

        if (cellStylerClass.isInterface()) {
            return false;
        }

        if (PxlConstants.VOID_CELL_STYLER.equals(cellStylerClass)) {
            return false;
        }

        return true;
    }

}
