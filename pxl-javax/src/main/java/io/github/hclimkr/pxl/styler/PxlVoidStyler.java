package io.github.hclimkr.pxl.styler;

import io.github.hclimkr.pxl.PxlConstants;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * No-op Excel cell styler.
 * <p>
 * Leaves the cell style untouched and applies no font.
 * <p>
 * It is <strong>not</strong> the "unset" marker: that is {@link PxlConstants#VOID_CELL_STYLER}, i.e.
 * {@link PxlStyler} itself, which the cascade recognizes through
 * {@code PxlMiscUtils.isEffectiveCellStylerClass} and skips to the next level. This class is a concrete
 * styler, so naming it ends the cascade and applies a blank style rather than inheriting the sheet's or
 * workbook's. PXL itself never selects it; it is kept for a caller that wants exactly that blank style.
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
