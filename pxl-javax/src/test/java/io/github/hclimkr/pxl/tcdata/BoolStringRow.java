package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying custom true/false string combinations (import/exportTrueString and FalseString).
 * For round-trip to work, the import and export strings must be set to the same values.
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
