package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.math.BigDecimal;
import java.math.BigInteger;

/**
 * DTO with numeric fields carrying an {@code exportMasking} regex (but no pattern), so each value is rendered
 * as text and masked on export, exercising the masking branch of each numeric codec's export-string method.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NumberMaskingRow {

    @PxlColumn(name = "WrapByte", exportMasking = "\\d")
    private Byte wrapByte;

    @PxlColumn(name = "PrimByte", exportMasking = "\\d")
    private byte primByte;

    @PxlColumn(name = "WrapShort", exportMasking = "\\d")
    private Short wrapShort;

    @PxlColumn(name = "PrimShort", exportMasking = "\\d")
    private short primShort;

    @PxlColumn(name = "WrapInt", exportMasking = "\\d")
    private Integer wrapInt;

    @PxlColumn(name = "PrimInt", exportMasking = "\\d")
    private int primInt;

    @PxlColumn(name = "WrapLong", exportMasking = "\\d")
    private Long wrapLong;

    @PxlColumn(name = "PrimLong", exportMasking = "\\d")
    private long primLong;

    @PxlColumn(name = "WrapDouble", exportMasking = "\\d")
    private Double wrapDouble;

    @PxlColumn(name = "PrimDouble", exportMasking = "\\d")
    private double primDouble;

    @PxlColumn(name = "WrapFloat", exportMasking = "\\d")
    private Float wrapFloat;

    @PxlColumn(name = "PrimFloat", exportMasking = "\\d")
    private float primFloat;

    @PxlColumn(name = "BigInt", exportMasking = "\\d")
    private BigInteger bigInt;

    @PxlColumn(name = "BigDec", exportMasking = "\\d")
    private BigDecimal bigDec;

}
