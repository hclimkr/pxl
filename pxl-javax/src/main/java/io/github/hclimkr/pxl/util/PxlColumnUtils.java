package io.github.hclimkr.pxl.util;

import io.github.hclimkr.pxl.PxlConstants;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Objects;

/**
 * Column-level POI helpers.
 * <p>
 * Currently a single entry point, {@code autoSizeColumns}, which widens a column to fit its content. It does not
 * simply delegate to POI: the computed width is scaled by 1.8 and then clamped between
 * {@link PxlConstants#EXPORT_AUTO_COLUMN_MIN_WIDTH} and {@link PxlConstants#EXPORT_AUTO_COLUMN_MAX_WIDTH}, because
 * POI's own measurement comes out too narrow for the fonts PXL exports with.
 * <p>
 * Auto-sizing measures every row of the column, so it is expensive on large sheets, and on a streaming (SXSSF)
 * sheet it only sees the rows still held in memory — the column must be tracked before those rows are flushed.
 */
public final class PxlColumnUtils {

    /**
     * Prevents instantiation.
     */
    private PxlColumnUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Auto-sizes a single column and then widens it heuristically: the POI-computed width is scaled by
     * 1.8 (capped at POI's maximum of 255 characters) and finally clamped between
     * {@link PxlConstants#EXPORT_AUTO_COLUMN_MIN_WIDTH} and
     * {@link PxlConstants#EXPORT_AUTO_COLUMN_MAX_WIDTH}. A {@code null} sheet or a negative column index
     * is a no-op.
     *
     * @param sheet       the sheet whose column is resized
     * @param columnIndex the zero-based index of the column to auto-size
     */
    public static void autoSizeColumns(final Sheet sheet, final int columnIndex) {

        if (Objects.isNull(sheet) || columnIndex < 0) {
            return;
        }

        sheet.autoSizeColumn(columnIndex);

        // The autoSize-computed width is too small (a Malgun Gothic font issue?), so artificially adjust it to 1.8x
        int columnWidth = Math.min(255 * 256, (int) (sheet.getColumnWidth(columnIndex) * 1.8));

        // minimum width limit
        columnWidth = Math.max(columnWidth, PxlConstants.EXPORT_AUTO_COLUMN_MIN_WIDTH);

        // maximum width limit
        columnWidth = Math.min(columnWidth, PxlConstants.EXPORT_AUTO_COLUMN_MAX_WIDTH);

        sheet.setColumnWidth(columnIndex, columnWidth);
    }

}
