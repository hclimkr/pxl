package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.Duration;
import java.time.LocalDate;
import java.time.Period;

/**
 * DTO for verifying that importing an invalid value per type throws an exception.
 * Each test supplies a fixture containing only the relevant column so that only that column's codec error is triggered.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TypedRow {

    @PxlColumn(name = "Bool")
    private Boolean bool;

    @PxlColumn(name = "Dec")
    private BigDecimal dec;

    @PxlColumn(name = "Int")
    private BigInteger bigInt;

    @PxlColumn(name = "Date")
    private LocalDate date;

    @PxlColumn(name = "Dur")
    private Duration dur;

    @PxlColumn(name = "Per")
    private Period per;

    @PxlColumn(name = "Num")
    private Integer num;

    @PxlColumn(name = "Small")
    private Byte small;

}
