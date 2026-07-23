package io.github.hclimkr.pxl.styler.header;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Header cell styler with horizontal center alignment.
 * <p>
 * Extends {@link PxlHeaderStyler} (bold font, vertical center, grey-25% fill) and additionally sets
 * {@link HorizontalAlignment#CENTER} on the cell style.
 */
public class PxlHeaderHorizontalCenterTextStyler extends PxlHeaderStyler {

    /**
     * Applies the base header formatting, then sets horizontal center alignment.
     *
     * @param workbook  the workbook, used to create the base bold {@link Font}
     * @param cellStyle the cell style to mutate in place
     * @return the {@link Font} produced by the base header styler
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        cellStyle.setAlignment(HorizontalAlignment.CENTER);

        return font;
    }

}
