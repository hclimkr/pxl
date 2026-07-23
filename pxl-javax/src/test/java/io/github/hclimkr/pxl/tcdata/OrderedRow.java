package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying exportOrder (column order).
 * <p>
 * The field declaration order is X, Y, Z, but according to exportOrder (alphabetical) the export order must be Y(A), Z(B), X(C).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class OrderedRow {

    @PxlColumn(name = "X", exportOrder = "C")
    private String x;

    @PxlColumn(name = "Y", exportOrder = "A")
    private String y;

    @PxlColumn(name = "Z", exportOrder = "B")
    private String z;

}
