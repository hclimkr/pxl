package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Verifies column-level exportSampleEnabled. The skip column must be excluded from the sample.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SampleColumnRow {

    @PxlColumn(name = "Keep", exportSample = "K")
    private String keep;

    @PxlColumn(name = "Skip", exportSample = "S", exportSampleEnabled = false)
    private String skip;

}
