package io.github.hclimkr.pxl.styler.data;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Data cell styler with horizontal center alignment.
 * <p>
 * Sets {@link HorizontalAlignment#CENTER} on the cell style. Adds no font of its own.
 */
public class PxlDataHorizontalCenterTextStyler extends PxlDataStyler {

    /**
     * Sets horizontal center alignment on the cell style.
     *
     * @param workbook  the workbook (unused)
     * @param cellStyle the cell style to mutate in place
     * @return {@code null}, inherited from the base data styler (no font is set)
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        cellStyle.setAlignment(HorizontalAlignment.CENTER);

        return font;
    }

}
