package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Date;

/**
 * DTO for verifying import of files filled with NUMERIC/BOOLEAN cells, as if produced externally (Excel/POI).
 * <p>
 * Covers the codec's NUMERIC/BOOLEAN branches that a round-trip (where PXL exports as STRING/text) never exercises.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NumericCellRow {

    // NUMERIC (date-formatted) cell -> each date/time type (DateUtil.isCellDateFormatted / getLocalDateTimeCellValue path)
    @PxlColumn(name = "LocalDate")
    private LocalDate localDate;

    @PxlColumn(name = "LocalTime")
    private LocalTime localTime;

    @PxlColumn(name = "LocalDateTime")
    private LocalDateTime localDateTime;

    @PxlColumn(name = "JavaDate")
    private Date javaDate;

    @PxlColumn(name = "ZonedDateTime")
    private ZonedDateTime zonedDateTime;

    @PxlColumn(name = "OffsetDateTime")
    private OffsetDateTime offsetDateTime;

    // NUMERIC cell -> numeric type
    @PxlColumn(name = "IntVal")
    private int intVal;

    @PxlColumn(name = "LongVal")
    private long longVal;

    @PxlColumn(name = "DoubleVal")
    private double doubleVal;

    @PxlColumn(name = "BigDec")
    private BigDecimal bigDec;

    @PxlColumn(name = "BigInt")
    private BigInteger bigInt;

    // NUMERIC cell -> String (DataFormatter render; e.g. "2012000046" without exponent notation)
    @PxlColumn(name = "NumericAsString")
    private String numericAsString;

    // NUMERIC cell -> char (first character after NumberToTextConverter)
    @PxlColumn(name = "CharFromNumeric")
    private char charFromNumeric;

    // Genuine BOOLEAN cell -> Boolean
    @PxlColumn(name = "BoolFromBoolean")
    private Boolean boolFromBoolean;

    // NUMERIC cell -> Boolean (non-zero)
    @PxlColumn(name = "BoolFromNumeric")
    private Boolean boolFromNumeric;

    // NUMERIC cell -> Duration (interpreted as seconds)
    @PxlColumn(name = "DurationFromNumeric")
    private Duration durationFromNumeric;

    // NUMERIC cell -> Period (interpreted as days)
    @PxlColumn(name = "PeriodFromNumeric")
    private Period periodFromNumeric;

}
