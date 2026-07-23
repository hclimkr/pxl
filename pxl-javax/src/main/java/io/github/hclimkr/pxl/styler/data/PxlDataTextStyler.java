package io.github.hclimkr.pxl.styler.data;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Data cell styler that forces text format.
 * <p>
 * Sets the cell data format to {@code "@"} so numeric-looking content is stored and displayed verbatim as text
 * (preserving leading zeros and preventing numeric coercion). Adds no font of its own.
 */
public class PxlDataTextStyler extends PxlDataStyler {

    /**
     * Applies the {@code "@"} (text) data format to the cell style.
     *
     * @param workbook  the workbook, used to create the {@code DataFormat}
     * @param cellStyle the cell style to mutate in place
     * @return {@code null}, inherited from the base data styler (no font is set)
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        final DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("@"));

        return font;
    }

}
