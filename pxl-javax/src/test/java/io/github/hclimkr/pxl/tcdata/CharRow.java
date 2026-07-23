package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * For verifying char/Character conversion from numeric cells and boolean cells.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CharRow {

    @PxlColumn(name = "C")
    private Character c;

}
