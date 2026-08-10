package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO whose every column opts out of the sample, leaving a sample export with no column to write.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NoSampleColumnRow {

    @PxlColumn(name = "First", exportSample = "F", exportSampleEnabled = false)
    private String first;

    @PxlColumn(name = "Second", exportSample = "S", exportSampleEnabled = false)
    private String second;

}
