package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO with primitive numeric fields carrying a DecimalFormat pattern.
 * A pattern makes each value exported as text and re-parsed via DecimalFormat on import,
 * exercising the primitive codec's exported-to-string / DecimalFormat branches on both directions.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class PrimitivePatternRow {

    @PxlColumn(name = "LongCount", pattern = "#,##0")
    private long longCount;

    @PxlColumn(name = "IntCount", pattern = "#,##0")
    private int intCount;

    @PxlColumn(name = "DoubleAmt", pattern = "#,##0.00")
    private double doubleAmt;

}
