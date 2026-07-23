package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Verifies that a String column renders a BOOLEAN cell using importTrueString/FalseString.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class StringFlagRow {

    @PxlColumn(name = "Flag", importTrueString = "YES", importFalseString = "NO")
    private String flag;

}
