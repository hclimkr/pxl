package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.styler.header.PxlHeaderHorizontalCenterTextStyler;
import lombok.*;

import javax.validation.constraints.NotNull;

/**
 * For verifying the per-column required header styler (exportColumnRequiredHeaderCellStyler).
 * <p>
 * Because it is @NotNull, the required header styler is selected, and the specified horizontal-center styler is applied to the header.
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

}
