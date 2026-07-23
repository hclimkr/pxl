package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.math.BigDecimal;
import java.math.BigInteger;

/**
 * DTO with wrapper and primitive numeric fields carrying a DecimalFormat pattern.
 * A pattern makes each value exported as text and re-parsed via DecimalFormat on import,
 * exercising each numeric codec's exported-to-string / DecimalFormat branch on both directions.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NumberPatternRow {

    @PxlColumn(name = "WrapByte", pattern = "#,##0")
    private Byte wrapByte;

    @PxlColumn(name = "PrimByte", pattern = "#,##0")
    private byte primByte;

    @PxlColumn(name = "WrapShort", pattern = "#,##0")
    private Short wrapShort;

    @PxlColumn(name = "PrimShort", pattern = "#,##0")
    private short primShort;

    @PxlColumn(name = "WrapInt", pattern = "#,##0")
    private Integer wrapInt;

    @PxlColumn(name = "PrimInt", pattern = "#,##0")
    private int primInt;

    @PxlColumn(name = "WrapLong", pattern = "#,##0")
    private Long wrapLong;

    @PxlColumn(name = "PrimLong", pattern = "#,##0")
    private long primLong;

    @PxlColumn(name = "WrapFloat", pattern = "#,##0.0")
    private Float wrapFloat;

    @PxlColumn(name = "PrimFloat", pattern = "#,##0.0")
    private float primFloat;

    @PxlColumn(name = "WrapDouble", pattern = "#,##0.00")
    private Double wrapDouble;

    @PxlColumn(name = "PrimDouble", pattern = "#,##0.00")
    private double primDouble;

    @PxlColumn(name = "BigInt", pattern = "#,##0")
    private BigInteger bigInt;

    @PxlColumn(name = "BigDec", pattern = "#,##0.00")
    private BigDecimal bigDec;

}
