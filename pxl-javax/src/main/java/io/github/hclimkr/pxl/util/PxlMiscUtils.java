package io.github.hclimkr.pxl.util;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.styler.PxlStyler;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.Objects;
import java.util.regex.Pattern;

/**
 * Miscellaneous spreadsheet helpers: conversions between numeric row/column indexes and A1-style
 * column letters, cell references and range addresses, a filter for text XML cannot carry, plus a
 * check for whether a styler class is an effective (usable) cell styler.
 * <p>
 * The conversions settle the shape of what they are given before POI sees it, in both directions: letters that are
 * not column letters and a negative index alike raise {@link PxlArgumentException}. Left to itself POI answers
 * some of these with a value rather than a refusal - {@code "1"} as column {@code -16}, column {@code -1} as an
 * empty string, row {@code -1} as a reference with no row in it - and a wrong reference is harder to notice than a
 * failure. Only the lower bound is checked; an index past a format's limit still converts, since these methods
 * build text and are not told which format it is for.
 */
public final class PxlMiscUtils {

    /**
     * The shape of A1-style column letters: the letters themselves, optionally behind an absolute-reference
     * marker ({@code $}).
     */
    private static final Pattern COLUMN_STRING_PATTERN = Pattern.compile("\\$?[A-Za-z]+");

    /**
     * Prevents instantiation.
     */
    private PxlMiscUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Converts a zero-based column index to its A1-style column letters (for example {@code 0 -> "A"},
     * {@code 26 -> "AA"}).
     * <p>
     * A negative index is refused rather than converted: POI answers one with an empty string, which reads as a
     * column reference right up until it is used as one. This is the mirror of
     * {@link #convertColumnStringToColumnIndex}, which likewise settles the shape of its argument before POI sees
     * it.
     *
     * @param columnIndex the zero-based column index
     * @return the column letters
     * @throws PxlArgumentException if {@code columnIndex} is negative
     */
    public static String convertColumnIndexToColumnString(final int columnIndex)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(columnIndex, "columnIndex");

        return CellReference.convertNumToColString(columnIndex);
    }

    /**
     * Converts A1-style column letters (for example {@code "A"}, {@code "AA"}) to a zero-based column index.
     * An absolute-reference marker ({@code $}) is accepted and ignored.
     *
     * @param columnString the column letters
     * @return the zero-based column index
     * @throws PxlNullPointerException if {@code columnString} is {@code null}
     * @throws PxlArgumentException    if {@code columnString} is blank or is not made of column letters
     */
    public static int convertColumnStringToColumnIndex(final String columnString)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notBlank(columnString, "columnString");

        // POI does not check that what it is given is column letters: it folds every character into the running
        // total the same way and answers a nonsense (often negative) index for anything else, so the shape is
        // settled here before it gets there.
        if (!COLUMN_STRING_PATTERN.matcher(columnString).matches()) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_COLUMN_STRING_INVALID, columnString));
        }

        return CellReference.convertColStringToIndex(columnString);
    }

    /**
     * Builds an A1-style cell reference (for example {@code "B3"}) from zero-based row/column indexes,
     * formatted without absolute-reference (`$`) markers.
     * <p>
     * A negative index is refused. POI reads {@code -1} as "not stated" rather than as an error and builds half a
     * reference from it - {@code (-1, 0)} comes back as {@code "A"} and {@code (0, -1)} as {@code "1"} - which is
     * a worse answer than none; anything below that it turns down itself, with a bare
     * {@link IllegalArgumentException}.
     *
     * @param rowIndex    the zero-based row index
     * @param columnIndex the zero-based column index
     * @return the A1-style cell reference
     * @throws PxlArgumentException if either index is negative
     */
    public static String convertIndexesToCellReferenceString(final int rowIndex,
                                                             final int columnIndex)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(rowIndex, "rowIndex");
        PxlAssertSupport.notNegative(columnIndex, "columnIndex");

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
     * @throws PxlArgumentException if any of the indexes is negative
     */
    public static String convertIndexesToCellRangeAddressString(final int startRowIndex,
                                                                int startColumnIndex,
                                                                int endRowIndex,
                                                                int endColumnIndex)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(startRowIndex, "startRowIndex");
        PxlAssertSupport.notNegative(startColumnIndex, "startColumnIndex");
        PxlAssertSupport.notNegative(endRowIndex, "endRowIndex");
        PxlAssertSupport.notNegative(endColumnIndex, "endColumnIndex");

        final CellRangeAddress cellRangeAddress = new CellRangeAddress(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);

        return cellRangeAddress.formatAsString();
    }

    /**
     * Parses an A1-style cell reference into a (row index, column index) pair, both zero-based. The
     * reference must carry both a row and a column; a column-only or row-only reference is rejected.
     *
     * @param cellRefStr the A1-style cell reference
     * @return a pair of (zero-based row index, zero-based column index)
     * @throws PxlNullPointerException if {@code cellRefStr} is {@code null}
     * @throws PxlArgumentException    if {@code cellRefStr} is blank, cannot be read as a cell reference at all,
     *                                 or does not include both a row and a column
     */
    public static Pair<Integer, Integer> convertCellReferenceStringToIndexes(final String cellRefStr)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notBlank(cellRefStr, "cellRefStr");

        final CellReference cellRef;
        try {
            cellRef = new CellReference(cellRefStr);
        } catch (IllegalArgumentException e) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_CELL_REF_INVALID, cellRefStr), e);
        }

        if (cellRef.getRow() < 0 || cellRef.getCol() < 0) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_CELL_REF_INVALID, cellRefStr));
        }

        return Pair.of(cellRef.getRow(), (int) cellRef.getCol());
    }

    /**
     * Determines whether the given text carries a character that XML cannot represent, and so cannot survive an
     * XLSX/XLSM export intact.
     * <p>
     * The set of characters XML 1.0 allows is {@code #x9}, {@code #xA}, {@code #xD},
     * {@code #x20-#xD7FF}, {@code #xE000-#xFFFD} and {@code #x10000-#x10FFFF}. Everything else - the C0 control
     * characters other than tab, newline and carriage return, {@code #xFFFE} and {@code #xFFFF}, and any surrogate
     * left without its pair - is invalid. The text is walked by code point, so a well-formed surrogate pair counts
     * as the supplementary character it encodes rather than as two invalid halves.
     * <p>
     * Nothing in the write path reports these characters: POI 5.5.1 turns each of them into {@code '?'} on the way
     * into an XLSX (XSSF through the XmlBeans saver, SXSSF through its own sheet writer) without an exception or a
     * warning, while XLS keeps them as they are because the binary format is not XML. The loss is therefore silent
     * and it splits by file format, which is what makes it worth testing for up front.
     *
     * @param value the text to test, may be {@code null}
     * @return {@code true} if the text carries at least one character XML cannot represent; {@code false} for
     * {@code null} or empty text
     */
    public static boolean containsInvalidXmlChars(final String value) {

        if (Objects.isNull(value)) {
            return false;
        }

        final int length = value.length();
        int index = 0;
        while (index < length) {
            final int codePoint = value.codePointAt(index);
            if (!isValidXmlCodePoint(codePoint)) {
                return true;
            }
            index += Character.charCount(codePoint);
        }

        return false;
    }

    /**
     * Removes every character that XML cannot represent, leaving the rest of the text untouched. The characters
     * removed are the ones {@link #containsInvalidXmlChars} reports on.
     * <p>
     * Removing is a deliberate choice over the {@code '?'} POI substitutes: a dropped character reads as an
     * absence, whereas a substituted one reads as content that was really there. Callers who want POI's
     * behavior can keep it by not filtering at all.
     * <p>
     * Nothing is copied when there is nothing to remove - the argument itself comes back, {@code null} included -
     * so this is cheap to call on text that is usually clean.
     *
     * @param value the text to filter, may be {@code null}
     * @return the text without the characters XML cannot represent; the argument itself when it carries none
     */
    public static String removeInvalidXmlChars(final String value) {

        if (!containsInvalidXmlChars(value)) {
            return value;
        }

        final int length = value.length();
        final StringBuilder builder = new StringBuilder(length);
        int index = 0;
        while (index < length) {
            final int codePoint = value.codePointAt(index);
            final int charCount = Character.charCount(codePoint);
            if (isValidXmlCodePoint(codePoint)) {
                builder.append(value, index, index + charCount);
            }
            index += charCount;
        }

        return builder.toString();
    }

    /**
     * Determines whether a single code point is one XML 1.0 allows in character data.
     *
     * @param codePoint the code point to test
     * @return {@code true} if XML can represent it
     */
    private static boolean isValidXmlCodePoint(final int codePoint) {

        return codePoint == 0x9
                || codePoint == 0xA
                || codePoint == 0xD
                || (codePoint >= 0x20 && codePoint <= 0xD7FF)
                || (codePoint >= 0xE000 && codePoint <= 0xFFFD)
                || (codePoint >= 0x10000 && codePoint <= 0x10FFFF);
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
