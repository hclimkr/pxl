package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * (Guard regression) Verifies that an invalid regex exportMasking leads to an exception at build time.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadMaskingRow {

    @PxlColumn(name = "Value", exportMasking = "[invalid(")   // uncompilable regex
    private String value;

}
