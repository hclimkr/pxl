package io.github.hclimkr.pxl.styler.header;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.styler.PxlStyler;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Base Excel header cell styler.
 * <p>
 * Applies a bold font, vertical center alignment, and a solid grey-25% foreground fill
 * ({@link PxlConstants#HEADER_COLUMN_FOREGROUND_COLOR} / {@link PxlConstants#HEADER_COLUMN_FILL_PATTERN}).
 * Subclasses build on this to add required/optional coloring, wrapping, or explicit alignment.
 */
public class PxlHeaderStyler implements PxlStyler {

    /**
     * Applies the base header formatting: bold font, vertical center alignment, and solid grey-25% fill.
     *
     * @param workbook  the workbook, used to create the bold {@link Font}
     * @param cellStyle the cell style to mutate in place
     * @return the created bold {@link Font} (never {@code null})
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = workbook.createFont();
        font.setBold(true);

        cellStyle.setFont(font);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setFillForegroundColor(PxlConstants.HEADER_COLUMN_FOREGROUND_COLOR.getIndex());
        cellStyle.setFillPattern(PxlConstants.HEADER_COLUMN_FILL_PATTERN);

        return font;
    }

}
