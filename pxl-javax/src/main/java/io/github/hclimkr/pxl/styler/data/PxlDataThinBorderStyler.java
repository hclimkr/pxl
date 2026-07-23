package io.github.hclimkr.pxl.styler.data;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Data cell styler that draws a thin border on all four sides.
 * <p>
 * Sets {@link BorderStyle#THIN} on the left, right, top, and bottom edges of the cell. Adds no font of its own.
 */
public class PxlDataThinBorderStyler extends PxlDataStyler {

    /**
     * Applies a thin border to all four edges of the cell style.
     *
     * @param workbook  the workbook (unused for border formatting)
     * @param cellStyle the cell style to mutate in place
     * @return {@code null}, inherited from the base data styler (no font is set)
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);

        return font;
    }

}
