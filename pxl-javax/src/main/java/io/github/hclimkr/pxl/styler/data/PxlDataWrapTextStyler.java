package io.github.hclimkr.pxl.styler.data;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Data cell styler that enables text wrapping.
 * <p>
 * Turns on wrap-text so long values break across multiple lines within the cell. Adds no font of its own.
 */
public class PxlDataWrapTextStyler extends PxlDataStyler {

    /**
     * Enables wrap-text on the cell style.
     *
     * @param workbook  the workbook (unused)
     * @param cellStyle the cell style to mutate in place
     * @return {@code null}, inherited from the base data styler (no font is set)
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        cellStyle.setWrapText(true);

        return font;
    }

}
