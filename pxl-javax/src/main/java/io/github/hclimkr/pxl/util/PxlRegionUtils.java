package io.github.hclimkr.pxl.util;

import com.github.pjfanning.xlsx.impl.StreamingSheet;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * Merged-region helpers: looking up the merged region covering a cell, replicating a source row's
 * merged regions onto other rows, and removing merged regions within a row range. Streaming sheets
 * ({@link StreamingSheet}) do not expose merged regions, so the methods that write are a no-op there
 * and {@code getMergedRegion} yields {@code null}, exactly as they do on a {@code null} sheet.
 * <p>
 * Replicating onto a range of rows carries only the regions that fit in a single row, since consecutive
 * destinations are one row apart and a taller region would overlap the copy made for the row before it;
 * replicating onto one named row has no such restriction.
 */
public final class PxlRegionUtils {

    private static final Logger LOGGER = LoggerFactory.getLogger(PxlRegionUtils.class);

    /**
     * Prevents instantiation.
     */
    private PxlRegionUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Returns the merged region that covers the given cell position, or {@code null} if the position is
     * not part of any merged region. A {@code null} sheet or a streaming sheet ({@link StreamingSheet})
     * yields {@code null}.
     *
     * @param sheet       the sheet to inspect
     * @param rowIndex    the zero-based row index
     * @param columnIndex the zero-based column index
     * @return the covering merged region, or {@code null} if none applies
     */
    public static CellRangeAddress getMergedRegion(final Sheet sheet,
                                                   final int rowIndex,
                                                   final int columnIndex) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return null;
        }

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            final CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
            if (cellRangeAddress.isInRange(rowIndex, columnIndex)) {
                return cellRangeAddress;
            }
        }

        return null;
    }

    /**
     * Replicates every merged region anchored on the source row (i.e. whose first row equals the source
     * row's index) onto the destination row, preserving each region's row span and column extent. A
     * {@code null} sheet, a streaming sheet ({@link StreamingSheet}) or a {@code null} source/destination
     * row makes this a no-op, as does a destination row that is the source row itself - re-creating a
     * region at the coordinates it already occupies would only overlap the original.
     *
     * @param sheet  the sheet to operate on
     * @param srcRow the row whose anchored merged regions are copied
     * @param dstRow the row onto which the merged regions are re-created
     */
    public static void copyMergedRegionsInRow(final Sheet sheet,
                                              final Row srcRow,
                                              final Row dstRow) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        if (ObjectUtils.anyNull(srcRow, dstRow)) {
            return;
        }

        if (srcRow.getRowNum() == dstRow.getRowNum()) {
            return;
        }

        sheet.getMergedRegions()
                .forEach(cellRangeAddress -> {
                    if (cellRangeAddress.getFirstRow() == srcRow.getRowNum()) {
                        final int firstRowOfMergedRegion = dstRow.getRowNum();
                        final int lastRowOfMergedRegion = firstRowOfMergedRegion + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow());

                        final CellRangeAddress newCellRangeAddress = cellRangeAddress.copy();
                        newCellRangeAddress.setFirstRow(firstRowOfMergedRegion);
                        newCellRangeAddress.setLastRow(lastRowOfMergedRegion);

                        sheet.addMergedRegion(newCellRangeAddress);
                    }
                });
    }

    /**
     * Replicates merged regions anchored on one row onto another, resolving both rows by index and
     * delegating to {@link #copyMergedRegionsInRow(Sheet, Row, Row)}. A {@code null} sheet or a
     * streaming sheet ({@link StreamingSheet}) makes this a no-op, as does naming the same row twice.
     *
     * @param sheet       the sheet to operate on
     * @param srcRowIndex the zero-based index of the source row
     * @param dstRowIndex the zero-based index of the destination row
     */
    public static void copyMergedRegionsInRow(final Sheet sheet,
                                              final int srcRowIndex,
                                              final int dstRowIndex) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        final Row srcRow = sheet.getRow(srcRowIndex);
        final Row dstRow = sheet.getRow(dstRowIndex);

        copyMergedRegionsInRow(sheet, srcRow, dstRow);
    }

    /**
     * Replicates the merged regions anchored on the source row onto every row in the inclusive range
     * {@code [dstStartRowIndex, dstEndRowIndex]}, resolving the source row by index and delegating to
     * {@link #copyMergedRegionsInRow(Sheet, Row, int, int)}. A {@code null} sheet or a streaming sheet
     * ({@link StreamingSheet}) makes this a no-op, and a range covering the source row skips that one row.
     * Merged regions taller than one row are skipped as well, as described on that method.
     *
     * @param sheet            the sheet to operate on
     * @param srcRowIndex      the zero-based index of the source row
     * @param dstStartRowIndex the zero-based index of the first destination row (inclusive)
     * @param dstEndRowIndex   the zero-based index of the last destination row (inclusive)
     */
    public static void copyMergedRegionsInRow(final Sheet sheet,
                                              final int srcRowIndex,
                                              final int dstStartRowIndex,
                                              final int dstEndRowIndex) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        copyMergedRegionsInRow(sheet, sheet.getRow(srcRowIndex), dstStartRowIndex, dstEndRowIndex);
    }

    /**
     * Replicates every merged region anchored on the source row (i.e. whose first row equals the source
     * row's index) onto each existing row in the inclusive range {@code [dstStartRowIndex, dstEndRowIndex]},
     * preserving each region's column extent. A {@code null} sheet, a streaming sheet
     * ({@link StreamingSheet}), a {@code null} source row, or an inverted range
     * ({@code dstStartRowIndex > dstEndRowIndex}) makes this a no-op; destination rows that do not exist
     * are skipped, as is the source row itself when the range covers it - re-creating a region at the
     * coordinates it already occupies would only overlap the original, and the rest of the range is
     * copied either way.
     * <p>
     * Only merged regions that fit in a single row (i.e. whose first and last row are the same) are
     * replicated. The destination rows sit one row apart, so a taller region would overlap the copy made
     * for the row before it, which POI rejects; such a region is skipped with a {@code WARN} log rather
     * than left to fail once part of the range is already merged. To carry a multi-row region over,
     * name a single destination with {@link #copyMergedRegionsInRow(Sheet, Row, Row)} instead.
     *
     * @param sheet            the sheet to operate on
     * @param srcRow           the row whose anchored merged regions are copied
     * @param dstStartRowIndex the zero-based index of the first destination row (inclusive)
     * @param dstEndRowIndex   the zero-based index of the last destination row (inclusive)
     */
    public static void copyMergedRegionsInRow(final Sheet sheet,
                                              final Row srcRow,
                                              final int dstStartRowIndex,
                                              final int dstEndRowIndex) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        if (Objects.isNull(srcRow)) {
            return;
        }

        if (dstStartRowIndex > dstEndRowIndex) {
            return;
        }

        sheet.getMergedRegions()
                .forEach(cellRangeAddress -> {
                    if (cellRangeAddress.getFirstRow() == srcRow.getRowNum()) {
                        if (cellRangeAddress.getFirstRow() != cellRangeAddress.getLastRow()) {
                            // The destination rows sit one row apart, so a region taller than one row would
                            // always overlap the copy made for the row before it - POI rejects the second one
                            // with an IllegalStateException. Skip such a region instead of failing halfway
                            // through, having already merged part of the range.
                            LOGGER.warn(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.LOG_MERGED_REGION_MULTI_ROW_SKIPPED,
                                    cellRangeAddress.formatAsString(),
                                    (dstStartRowIndex + 1) + ":" + (dstEndRowIndex + 1)));
                            return;
                        }

                        IntStream.range(dstStartRowIndex, dstEndRowIndex + 1)
                                .forEach(dstRowIndex -> {
                                    final Row dstRow = sheet.getRow(dstRowIndex);
                                    if (Objects.isNull(dstRow)) {
                                        return;
                                    }

                                    if (dstRow.getRowNum() == srcRow.getRowNum()) {
                                        return;
                                    }

                                    final int firstRowOfMergedRegion = dstRow.getRowNum();
                                    final int lastRowOfMergedRegion = firstRowOfMergedRegion + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow());

                                    final CellRangeAddress newCellRangeAddress = cellRangeAddress.copy();
                                    newCellRangeAddress.setFirstRow(firstRowOfMergedRegion);
                                    newCellRangeAddress.setLastRow(lastRowOfMergedRegion);

                                    sheet.addMergedRegion(newCellRangeAddress);
                                });
                    }
                });
    }

    /**
     * Removes every merged region wholly contained within the inclusive row range
     * {@code [startRowIndex, endRowIndex]} (i.e. whose first row is {@code >= startRowIndex} and whose last
     * row is {@code <= endRowIndex}). Regions that only partially overlap the range are left intact. A
     * {@code null} sheet, a streaming sheet ({@link StreamingSheet}), or an inverted range
     * ({@code startRowIndex > endRowIndex}) makes this a no-op.
     *
     * @param sheet         the sheet to operate on
     * @param startRowIndex the zero-based index of the first row of the range (inclusive)
     * @param endRowIndex   the zero-based index of the last row of the range (inclusive)
     */
    public static void removeMergedRegionInRows(final Sheet sheet,
                                                final int startRowIndex,
                                                final int endRowIndex) {

        if (Objects.isNull(sheet) || sheet instanceof StreamingSheet) {
            return;
        }

        if (startRowIndex > endRowIndex) {
            return;
        }

        final int numMergedRegions = sheet.getNumMergedRegions();
        final List<Integer> mergedRegionIndexList = IntStream.range(0, numMergedRegions)
                .filter(i -> {
                    final CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
                    return cellRangeAddress.getFirstRow() >= startRowIndex
                            && cellRangeAddress.getLastRow() <= endRowIndex;
                })
                .boxed()
                .collect(Collectors.toList());

        if (PxlCollectionUtils.isNotEmpty(mergedRegionIndexList)) {
            sheet.removeMergedRegions(mergedRegionIndexList);
        }
    }

}
