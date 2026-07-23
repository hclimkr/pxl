package io.github.hclimkr.pxl.styler.header;

import io.github.hclimkr.pxl.PxlConstants;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Objects;

/**
 * Header cell styler for optional columns.
 * <p>
 * Extends {@link PxlHeaderStyler} (bold font, vertical center, grey-25% fill) and recolors the font grey-50%
 * ({@link PxlConstants#OPTIONAL_HEADER_COLUMN_FONT_COLOR}) to visually distinguish optional columns, reapplying
 * the optional foreground fill.
 */
public class PxlHeaderOptionalStyler extends PxlHeaderStyler {

    /**
     * Applies the base header formatting, then sets the font color to grey-50% and reapplies the optional fill.
     *
     * @param workbook  the workbook, used to create a bold {@link Font} if the superclass returned none
     * @param cellStyle the cell style to mutate in place
     * @return the {@link Font} whose color was set to grey-50% (never {@code null})
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        Font font = super.apply(workbook, cellStyle);

        if (Objects.isNull(font)) {
            font = workbook.createFont();
            font.setBold(true);
            cellStyle.setFont(font);
        }

        font.setColor(PxlConstants.OPTIONAL_HEADER_COLUMN_FONT_COLOR.getIndex());

        cellStyle.setFillForegroundColor(PxlConstants.OPTIONAL_HEADER_COLUMN_FOREGROUND_COLOR.getIndex());
        cellStyle.setFillPattern(PxlConstants.OPTIONAL_HEADER_COLUMN_FILL_PATTERN);

        return font;
    }

}
