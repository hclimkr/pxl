package io.github.hclimkr.pxl.styler.data;

import io.github.hclimkr.pxl.styler.PxlStyler;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Base data (body) cell styler.
 * <p>
 * Applies no formatting on its own and serves as the superclass for the concrete data stylers (borders, wrapping,
 * alignment, number/text formats), each of which calls back into this base before adding its own formatting.
 */
public class PxlDataStyler implements PxlStyler {

    /**
     * Applies no formatting.
     *
     * @param workbook  the workbook (unused)
     * @param cellStyle the cell style (left unchanged)
     * @return {@code null}, since no font is applied
     */
    @Override
    public Font apply(final Workbook workbook, final CellStyle cellStyle) {

        // No Applied Style

        return null;
    }

}
