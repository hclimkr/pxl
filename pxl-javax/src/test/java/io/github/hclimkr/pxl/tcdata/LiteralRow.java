package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Column that does not specify exportStringAsFormula (defaults to false).
 * A string starting with "=..." must be recorded/restored as literal text rather than evaluated as a formula.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class LiteralRow {

    @PxlColumn(name = "Expr")
    private String expr;

}
