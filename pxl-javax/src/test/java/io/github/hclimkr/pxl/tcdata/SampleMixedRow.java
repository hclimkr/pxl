package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Verifies sample rendering when a column declares no exportSample.
 * The filled column gets its exportSample value; the column without exportSample renders a blank sample cell.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SampleMixedRow {

    @PxlColumn(name = "Filled", exportSample = "V")
    private String filled;

    @PxlColumn(name = "Empty")
    private String empty;

}
