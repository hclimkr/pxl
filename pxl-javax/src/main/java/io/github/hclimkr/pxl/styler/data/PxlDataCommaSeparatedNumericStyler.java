package io.github.hclimkr.pxl.styler.data;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Data cell styler that applies thousands grouping to numbers.
 * <p>
 * Sets the cell data format to {@code "#,##0"}, so numeric values render with comma thousands separators and no
 * decimal places. Adds no font of its own.
 */
public class PxlDataCommaSeparatedNumericStyler extends PxlDataStyler {

    /**
     * Applies the {@code "#,##0"} (thousands-grouped) data format to the cell style.
     *
     * @param workbook  the workbook, used to create the {@code DataFormat}
     * @param cellStyle the cell style to mutate in place
     * @return {@code null}, inherited from the base data styler (no font is set)
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        final Font font = super.apply(workbook, cellStyle);

        final DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("#,##0"));

        return font;
    }

}
