package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * A DTO with a column (exportStringAsFormula) that exports a string as a formula.
 * On export it is written as a formula cell, and on (non-streaming) import the computed result is read as a string.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class FormulaRow {

    @PxlColumn(name = "Label")
    private String label;

    @PxlColumn(name = "Formula", exportStringAsFormula = true)
    private String formula;

}
