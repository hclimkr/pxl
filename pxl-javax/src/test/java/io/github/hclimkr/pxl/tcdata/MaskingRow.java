package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO with a column that has a masking rule (exportMasking) specified.
 * On export, characters matching the regex are replaced with '*', and that masked value round-trips as-is.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class MaskingRow {

    // Mask all digits.
    @PxlColumn(name = "Secret", exportMasking = "\\d")
    private String secret;

}
