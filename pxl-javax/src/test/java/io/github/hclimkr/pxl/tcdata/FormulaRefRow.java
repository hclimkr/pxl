package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * A DTO for verifying formulas that reference other cells (exportStringAsFormula).
 * <p>
 * exportOrder places the columns deterministically in columns A/B/C - Qty=column A, Price=column B, Total=column C.
 * Total holds the formula {@code =A2*B2} that multiplies Qty and Price of the same data row (Excel row 2),
 * to confirm that the cell-reference formula is actually computed (cached) at export time.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class FormulaRefRow {

    @PxlColumn(name = "Qty", exportOrder = "1")
    private Integer qty;

    @PxlColumn(name = "Price", exportOrder = "2")
    private Integer price;

    @PxlColumn(name = "Total", exportOrder = "3", exportStringAsFormula = true)
    private String total;

}
