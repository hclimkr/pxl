package io.github.hclimkr.pxl.styler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel cell styler strategy.
 * <p>
 * Implementations mutate the supplied {@link CellStyle} in place (borders, alignment, fill, data format, etc.)
 * and, when they need to change font attributes, create a {@link Font} on the workbook and attach it to the style.
 * Header stylers live under {@code styler.header} and data stylers under {@code styler.data}. Resolution cascades
 * column &rarr; sheet &rarr; workbook &rarr; built-in default.
 */
public interface PxlStyler {

    /**
     * Applies this styler's formatting to the given cell style.
     *
     * @param workbook  the workbook, used to create any {@link Font} or {@link DataFormat} the styler needs
     * @param cellStyle the cell style to mutate in place
     * @return the {@link Font} attached to the cell style, or {@code null} if this styler sets no font
     */
    Font apply(final Workbook workbook, final CellStyle cellStyle);

}
