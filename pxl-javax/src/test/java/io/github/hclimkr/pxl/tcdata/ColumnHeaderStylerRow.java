package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.styler.header.PxlHeaderHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderWrapTextStyler;
import lombok.*;

import javax.validation.constraints.NotNull;

/**
 * For verifying the per-column header stylers (exportColumnRequiredHeaderCellStyler / exportColumnOptionalHeaderCellStyler).
 * <p>
 * Which of the two is consulted follows from the column itself: custom is @NotNull, so it takes the required styler
 * (horizontal center), while plain has no such constraint and takes the optional one (wrap text).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ColumnHeaderStylerRow {

    @NotNull
    @PxlColumn(name = "Custom", exportColumnRequiredHeaderCellStyler = PxlHeaderHorizontalCenterTextStyler.class)
    private String custom;

    @PxlColumn(name = "Plain", exportColumnOptionalHeaderCellStyler = PxlHeaderWrapTextStyler.class)
    private String plain;

}
