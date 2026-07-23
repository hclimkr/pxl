package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.styler.data.PxlDataHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataThinBorderStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataWrapTextStyler;
import lombok.*;

import javax.validation.constraints.NotNull;

/**
 * DTO for verifying visual/dimensional attributes such as cell stylers and column width.
 * required is @NotNull so the Required header style applies; optional gets the Optional header style.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class StyledRow {

    @NotNull
    @PxlColumn(name = "Required")
    private String required;

    @PxlColumn(name = "Optional")
    private String optional;

    @PxlColumn(name = "Centered", exportColumnDataCellStyler = PxlDataHorizontalCenterTextStyler.class)
    private String centered;

    @PxlColumn(name = "Wrapped", exportColumnDataCellStyler = PxlDataWrapTextStyler.class)
    private String wrapped;

    @PxlColumn(name = "Bordered", exportColumnDataCellStyler = PxlDataThinBorderStyler.class)
    private String bordered;

    // Fixed column width (in units of 1/256 of a character width)
    @PxlColumn(name = "Wide", exportColumnWidth = 5000)
    private String wide;

}
