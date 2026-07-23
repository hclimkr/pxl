package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Verifies that enum matching is case- and whitespace-insensitive.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class EnumCaseRow {

    @PxlColumn(name = "Cat")
    private Category cat;

}
