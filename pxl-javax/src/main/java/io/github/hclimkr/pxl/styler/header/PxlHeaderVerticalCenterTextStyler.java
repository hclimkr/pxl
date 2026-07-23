package io.github.hclimkr.pxl.styler.header;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Header cell styler with vertical center alignment.
 * <p>
 * Extends {@link PxlHeaderStyler} (bold font, grey-25% fill) and explicitly sets
 * {@link VerticalAlignment#CENTER} on the cell style.
 */
public class PxlHeaderVerticalCenterTextStyler extends PxlHeaderStyler {

    /**
     * Applies the base header formatting, then sets vertical center alignment.
     *
     * @param workbook  the workbook, used to create the base bold {@link Font}
     * @param cellStyle the cell style to mutate in place
     * @return the {@link Font} produced by the base header styler
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        return font;
    }

}
