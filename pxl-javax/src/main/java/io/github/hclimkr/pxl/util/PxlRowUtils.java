package io.github.hclimkr.pxl.util;

import com.github.pjfanning.xlsx.impl.StreamingSheet;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Objects;
import java.util.Optional;
import java.util.stream.IntStream;

/**
 * Row-level POI helpers: blank-row detection, fetching a row (optionally creating it), copying a row - once, over a
 * destination range, or a number of times - and removing rows by range or count.
 * <p>
 * A copy carries the source row's height, its row style when it has one, every cell (through
 * {@link PxlCellUtils#copyCell}) and the merged regions anchored on it, shifting existing rows down when the
 * destination is occupied. A removal closes the gap by shifting the rows below up, dropping the merged regions
 * contained in the removed range first, so the sheet stays contiguous.
 * <p>
 * All of it rewrites sheet structure, which a streaming sheet ({@link StreamingSheet}) cannot do - every
 * sheet-taking method no-ops there, as it does on a {@code null} sheet or a missing source row. The count-based
 * overloads add nothing of their own beyond rejecting a non-positive count, and {@code isBlankRow} takes a row
 * rather than a sheet, so neither is affected.
 * <p>
 * A negative row index is refused with {@link PxlArgumentException} by everything that writes, since carrying it
 * through would only reach POI as a bare {@link IllegalArgumentException}; {@code getRow} refuses one only when it
 * is asked to create the row, and otherwise answers it with {@code null}, the way it answers an absent row. Only
 * the lower bound is checked - an index past the format's last row is left to POI, which names the limit.
 */
public final class PxlRowUtils {

    /**
     * Prevents instantiation.
     */
    private PxlRowUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Determines whether the row is blank, i.e. every cell up to its last cell is blank or holds only
     * whitespace text. A {@code null} row is treated as blank.
     *
     * @param row the row to inspect, may be {@code null}
     * @return {@code true} if the row is {@code null} or contains no non-blank cell value
     */
    public static boolean isBlankRow(final Row row) {

        if (Objects.isNull(row)) {
            return true;
        }

        for (int i = 0; i < row.getLastCellNum(); i++) {
            final Cell cell = row.getCell(i);
            if (!PxlCellUtils.isBlankCell(cell)) {
                final String cellValue = PxlCellUtils.getCellStringValue(cell);
                if (StringUtils.isNotBlank(cellValue)) {
                    return false;
                }
            }
        }

        return true;
    }

    /**
     * Returns the row at the given index, optionally creating it when absent. A {@code null} sheet or a
     * streaming sheet ({@link StreamingSheet}) yields {@code null}. When the row is missing and
     * {@code createIfNone} is {@code false}, {@code null} is returned.
     * <p>
     * A negative index is answered according to what is being asked for: a read yields {@code null}, the way an
     * absent row does, while creating a row at one is a change to the sheet and is refused with
     * {@link PxlArgumentException}. The refusal does not depend on the sheet, so an index that could never work
     * is turned down even when the sheet is {@code null} or streaming.
     *
     * @param sheet        the sheet to read from
     * @param rowIndex     the zero-based row index
     * @param createIfNone whether to create the row if it does not yet exist
     * @return the (possibly newly created) row, or {@code null} if unavailable
     * @throws PxlArgumentException if {@code rowIndex} is negative and {@code createIfNone} is {@code true}
     */
    public static Row getRow(final Sheet sheet,
                             final int rowIndex,
                             final boolean createIfNone)
            throws PxlArgumentException {

        if (rowIndex < 0) {
            if (createIfNone) {
                PxlAssertSupport.notNegative(rowIndex, "rowIndex");
            }

            return null;
        }

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return null;
        }

        Row row = sheet.getRow(rowIndex);
        if (Objects.isNull(row) && createIfNone) {
            row = sheet.createRow(rowIndex);
        }

        return row;
    }

    /**
     * Copies a row to another position within the same sheet, including height, the row style when the
     * source has one, every cell (via {@link PxlCellUtils#copyCell}) and any merged regions anchored on
     * the source row. If the destination row already exists, existing rows from it downward are first
     * shifted down by one to make room, and the source is then read at wherever that shift left it. A
     * {@code null} sheet, a streaming sheet ({@link StreamingSheet}) or a missing source row makes this
     * a no-op, as does giving the same index for source and destination.
     *
     * @param sheet       the sheet to operate on
     * @param srcRowIndex the zero-based index of the source row
     * @param dstRowIndex the zero-based index of the destination row
     * @throws PxlArgumentException if either index is negative
     */
    public static void copyRow(final Sheet sheet,
                               final int srcRowIndex,
                               final int dstRowIndex)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(srcRowIndex, "srcRowIndex");
        PxlAssertSupport.notNegative(dstRowIndex, "dstRowIndex");

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        if (dstRowIndex == srcRowIndex) {
            return;
        }

        if (Objects.isNull(sheet.getRow(srcRowIndex))) {
            return;
        }

        int shiftedSrcRowIndex = srcRowIndex;
        if (dstRowIndex <= sheet.getLastRowNum()) {
            // If destination row exist, push down all rows by 1
            sheet.shiftRows(dstRowIndex, sheet.getLastRowNum(), 1);

            if (srcRowIndex >= dstRowIndex) {
                shiftedSrcRowIndex = srcRowIndex + 1;
            }
        }

        // Read the source at its post-shift index rather than holding the reference taken before the shift:
        // XSSF moves the Row object itself and renumbers it, while HSSF moves the cells into another Row
        // object and empties this one, so only a re-read reference carries cells on both.
        final Row srcRow = sheet.getRow(shiftedSrcRowIndex);
        if (Objects.isNull(srcRow)) {
            return;
        }

        final Row dstRow = Optional.ofNullable(sheet.getRow(dstRowIndex))
                .orElseGet(() -> sheet.createRow(dstRowIndex));

        dstRow.setHeight(srcRow.getHeight());
        Optional.ofNullable(srcRow.getRowStyle()).ifPresent(dstRow::setRowStyle);

        IntStream.range(0, srcRow.getLastCellNum())
                .forEach(columnIndex -> {
                    final Cell srcCell = srcRow.getCell(columnIndex);
                    final Cell dstCell = dstRow.createCell(columnIndex);

                    PxlCellUtils.copyCell(srcCell, dstCell);
                });

        PxlRegionUtils.copyMergedRegionsInRow(sheet, srcRow, dstRow);
    }

    /**
     * Replicates a source row across an inclusive destination row range, copying height, the row style
     * when the source has one, and every cell into each target row, then re-applying the source row's
     * merged regions across the range. If the range start already exists, existing rows are first
     * shifted down to make room for the whole range, and the source is then read at wherever that shift
     * left it. A {@code null} sheet, a streaming sheet ({@link StreamingSheet}), an inverted range
     * ({@code dstStartRowIndex > dstEndRowIndex}) or a missing source row makes this a no-op, and a
     * destination row that is the source row itself once the shift has been applied is skipped. Of the
     * source row's merged regions only those fitting in a single row are replicated - see
     * {@link PxlRegionUtils#copyMergedRegionsInRow(Sheet, Row, int, int)}.
     *
     * @param sheet            the sheet to operate on
     * @param srcRowIndex      the zero-based index of the source row
     * @param dstStartRowIndex the zero-based first destination row (inclusive)
     * @param dstEndRowIndex   the zero-based last destination row (inclusive)
     * @throws PxlArgumentException if any of the indexes is negative
     */
    public static void copyRowMultiplyByRange(final Sheet sheet,
                                              final int srcRowIndex,
                                              final int dstStartRowIndex,
                                              final int dstEndRowIndex)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(srcRowIndex, "srcRowIndex");
        PxlAssertSupport.notNegative(dstStartRowIndex, "dstStartRowIndex");
        PxlAssertSupport.notNegative(dstEndRowIndex, "dstEndRowIndex");

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }
        if (dstStartRowIndex > dstEndRowIndex) {
            return;
        }

        if (Objects.isNull(sheet.getRow(srcRowIndex))) {
            return;
        }

        int shiftedSrcRowIndex = srcRowIndex;
        if (dstStartRowIndex <= sheet.getLastRowNum()) {
            // If destination rows exist, push down all rows
            final int shiftedRowCount = dstEndRowIndex - dstStartRowIndex + 1;
            sheet.shiftRows(dstStartRowIndex, sheet.getLastRowNum(), shiftedRowCount);

            if (srcRowIndex >= dstStartRowIndex) {
                shiftedSrcRowIndex = srcRowIndex + shiftedRowCount;
            }
        }

        // Read the source at its post-shift index rather than holding the reference taken before the shift:
        // XSSF moves the Row object itself and renumbers it, while HSSF moves the cells into another Row
        // object and empties this one, so only a re-read reference carries cells on both.
        final Row srcRow = sheet.getRow(shiftedSrcRowIndex);
        if (Objects.isNull(srcRow)) {
            return;
        }

        IntStream.range(dstStartRowIndex, dstEndRowIndex + 1)
                .forEach(dstRowIndex -> {
                    final Row dstRow = Optional.ofNullable(sheet.getRow(dstRowIndex))
                            .orElseGet(() -> sheet.createRow(dstRowIndex));

                    if (dstRow.getRowNum() == srcRow.getRowNum()) {
                        return;
                    }

                    dstRow.setHeight(srcRow.getHeight());
                    Optional.ofNullable(srcRow.getRowStyle()).ifPresent(dstRow::setRowStyle);

                    IntStream.range(0, srcRow.getLastCellNum())
                            .forEach(columnIndex -> {
                                final Cell srcCell = srcRow.getCell(columnIndex);
                                final Cell dstCell = dstRow.createCell(columnIndex);

                                PxlCellUtils.copyCell(srcCell, dstCell);
                            });
                });

        // srcRow was read at its post-shift index, so its row number matches where the regions moved to.
        PxlRegionUtils.copyMergedRegionsInRow(sheet, srcRow, dstStartRowIndex, dstEndRowIndex);
    }

    /**
     * Replicates a source row into {@code rowCount} consecutive rows starting at the given index, by
     * delegating to {@link #copyRowMultiplyByRange}. A non-positive {@code rowCount} is a no-op, but the
     * indexes are checked before that, so an unusable one is rejected either way.
     *
     * @param sheet            the sheet to operate on
     * @param srcRowIndex      the zero-based index of the source row
     * @param dstStartRowIndex the zero-based index of the first destination row
     * @param rowCount         the number of destination rows to create
     * @throws PxlArgumentException if either index is negative
     */
    public static void copyRowMultiplyByCount(final Sheet sheet,
                                              final int srcRowIndex,
                                              final int dstStartRowIndex,
                                              final int rowCount)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(srcRowIndex, "srcRowIndex");
        PxlAssertSupport.notNegative(dstStartRowIndex, "dstStartRowIndex");

        if (rowCount > 0) {
            copyRowMultiplyByRange(sheet, srcRowIndex, dstStartRowIndex, dstStartRowIndex + rowCount - 1);
        }
    }

    /**
     * Removes an inclusive range of rows and closes the gap by shifting the rows below up. Merged
     * regions fully contained in the range are removed first. A {@code null} sheet, a streaming sheet
     * ({@link StreamingSheet}), an inverted range ({@code startRowIndex > endRowIndex}) or an empty
     * sheet (no rows) makes this a no-op.
     *
     * @param sheet         the sheet to operate on
     * @param startRowIndex the zero-based first row to remove (inclusive)
     * @param endRowIndex   the zero-based last row to remove (inclusive)
     * @throws PxlArgumentException if either index is negative
     */
    public static void removeRowsByRange(final Sheet sheet,
                                         final int startRowIndex,
                                         final int endRowIndex)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(startRowIndex, "startRowIndex");
        PxlAssertSupport.notNegative(endRowIndex, "endRowIndex");

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }
        if (startRowIndex > endRowIndex) {
            return;
        }

        final int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return;
        }

        PxlRegionUtils.removeMergedRegionInRows(sheet, startRowIndex, endRowIndex);
        IntStream.range(startRowIndex, endRowIndex + 1)
                .forEach(rowIndex -> {
                    final Row removingRow = sheet.getRow(rowIndex);
                    if (Objects.nonNull(removingRow)) {
                        sheet.removeRow(removingRow);
                    }
                });

        if (endRowIndex < lastRowNum) {
            sheet.shiftRows(endRowIndex + 1, lastRowNum, startRowIndex - endRowIndex - 1);
        }
    }

    /**
     * Removes {@code rowCount} consecutive rows starting at the given index, by delegating to
     * {@link #removeRowsByRange}. A non-positive {@code rowCount} is a no-op, but the index is checked
     * before that, so an unusable one is rejected either way.
     *
     * @param sheet         the sheet to operate on
     * @param startRowIndex the zero-based index of the first row to remove
     * @param rowCount      the number of rows to remove
     * @throws PxlArgumentException if {@code startRowIndex} is negative
     */
    public static void removeRowsByCount(final Sheet sheet,
                                         final int startRowIndex,
                                         final int rowCount)
            throws PxlArgumentException {

        PxlAssertSupport.notNegative(startRowIndex, "startRowIndex");

        if (rowCount > 0) {
            removeRowsByRange(sheet, startRowIndex, startRowIndex + rowCount - 1);
        }
    }

}
