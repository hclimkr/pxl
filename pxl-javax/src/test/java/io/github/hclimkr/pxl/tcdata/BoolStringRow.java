package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying custom true/false string combinations (export/importTrueString and FalseString).
 * For round-trip to work, the export and import strings must be set to the same values.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BoolStringRow {

    @PxlColumn(name = "Flag",
            exportTrueString = "Y", exportFalseString = "N",
            importTrueString = "Y", importFalseString = "N")
    private Boolean flag;

}
