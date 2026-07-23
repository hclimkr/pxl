package io.github.hclimkr.pxl.styler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * No-op Excel cell styler.
 * <p>
 * Leaves the cell style untouched and applies no font. Used as the "unset" marker so cascade resolution can
 * skip to the next level.
 */
public class PxlVoidStyler implements PxlStyler {

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
