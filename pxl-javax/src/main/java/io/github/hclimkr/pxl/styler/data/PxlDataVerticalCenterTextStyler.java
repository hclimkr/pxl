package io.github.hclimkr.pxl.styler.data;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Data cell styler with vertical center alignment.
 * <p>
 * Sets {@link VerticalAlignment#CENTER} on the cell style. This is the built-in default data styler for exported
 * workbooks ({@code PxlConstants.DEFAULT_EXPORT_WORKBOOK_DATA_CELL_STYLER}). Adds no font of its own.
 */
public class PxlDataVerticalCenterTextStyler extends PxlDataStyler {

    /**
     * Sets vertical center alignment on the cell style.
     *
     * @param workbook  the workbook (unused)
     * @param cellStyle the cell style to mutate in place
     * @return {@code null}, inherited from the base data styler (no font is set)
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        return font;
    }

}
