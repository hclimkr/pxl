package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Date;
import java.util.List;

/**
 * DTO holding all scalar types supported by PXL plus Collection/enum/custom objects in a single row.
 * <p>
 * For Excel export -> import roundtrip verification. To allow roundtripping with the default pattern (no pattern specified),
 * all date/time values use only "second" granularity (the default pattern does not preserve nanoseconds/offset/zone).
 * <p>
 * Each column's exportSample is the value filled into the example row during Sample (blank form) export, and must be a string parseable as that type.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class AllTypesRow {

    // Field into which the data row's index is injected during import (not an export target)
    @PxlRowIndex
    private Integer rowIndex;

    // ----- String -----

    @PxlColumn(name = "Text", exportSample = "Sample text")
    private String text;

    // Verify that a number-looking string (leading zero) is preserved as a string
    @PxlColumn(name = "LeadingZero", exportSample = "007")
    private String leadingZero;

    // ----- Integer types (primitive/wrapper) -----

    @PxlColumn(name = "PrimByte", exportSample = "1")
    private byte primByte;

    @PxlColumn(name = "WrapByte", exportSample = "2")
    private Byte wrapByte;

    @PxlColumn(name = "PrimShort", exportSample = "3")
    private short primShort;

    @PxlColumn(name = "WrapShort", exportSample = "4")
    private Short wrapShort;

    @PxlColumn(name = "PrimInt", exportSample = "5")
    private int primInt;

    @PxlColumn(name = "WrapInt", exportSample = "6")
    private Integer wrapInt;

    @PxlColumn(name = "PrimLong", exportSample = "7")
    private long primLong;

    @PxlColumn(name = "WrapLong", exportSample = "8")
    private Long wrapLong;

    // ----- Floating-point types (primitive/wrapper) -----

    @PxlColumn(name = "PrimDouble", exportSample = "1.5")
    private double primDouble;

    @PxlColumn(name = "WrapDouble", exportSample = "2.5")
    private Double wrapDouble;

    @PxlColumn(name = "PrimFloat", exportSample = "3.5")
    private float primFloat;

    @PxlColumn(name = "WrapFloat", exportSample = "4.5")
    private Float wrapFloat;

    // ----- Character (primitive/wrapper) -----

    @PxlColumn(name = "PrimChar", exportSample = "A")
    private char primChar;

    @PxlColumn(name = "WrapChar", exportSample = "B")
    private Character wrapChar;

    // ----- Boolean (primitive/wrapper) -----

    @PxlColumn(name = "PrimBool", exportSample = "true")
    private boolean primBool;

    @PxlColumn(name = "WrapBool", exportSample = "false")
    private Boolean wrapBool;

    // ----- Big numbers (exported as text cells to preserve precision) -----

    @PxlColumn(name = "BigInt", exportSample = "12345678901234567890")
    private BigInteger bigInt;

    @PxlColumn(name = "BigDec", exportSample = "12345.6789")
    private BigDecimal bigDec;

    // ----- Date/time (second granularity) -----

    // java.util.Date roundtrips down to seconds because the default read pattern is ordered with second-inclusive priority. (no pattern specified)
    @PxlColumn(name = "JavaDate", pattern = "yyyy-MM-dd HH:mm:ss", exportSample = "2023-06-15 10:30:45")
    private Date javaDate;

    @PxlColumn(name = "LocalDate", pattern = "yyyy-MM-dd", exportSample = "2023-06-15")
    private LocalDate localDate;

    @PxlColumn(name = "LocalTime", pattern = "HH:mm:ss", exportSample = "10:30:45")
    private LocalTime localTime;

    @PxlColumn(name = "LocalDateTime", pattern = "yyyy-MM-dd HH:mm:ss", exportSample = "2023-06-15 10:30:45")
    private LocalDateTime localDateTime;

    @PxlColumn(name = "ZonedDateTime", exportSample = "2023-06-15T10:30:45+09:00")
    private ZonedDateTime zonedDateTime;

    @PxlColumn(name = "OffsetTime", exportSample = "10:30:45+09:00")
    private OffsetTime offsetTime;

    @PxlColumn(name = "OffsetDateTime", exportSample = "2023-06-15T10:30:45+09:00")
    private OffsetDateTime offsetDateTime;

    @PxlColumn(name = "Duration", exportSample = "PT1H2M3S")
    private Duration duration;

    @PxlColumn(name = "Period", exportSample = "P1Y2M3D")
    private Period period;

    // ----- enum -----

    @PxlColumn(name = "Grade", exportSample = "A")
    private Grade grade;

    @PxlColumn(name = "Category", exportSample = "Electronics")
    private Category category;

    // ----- Custom objects -----

    @PxlColumn(name = "Point", exportSample = "3,7")
    private Point point;

    @PxlColumn(name = "Money", exportSample = "USD 1050")
    private Money money;

    // ----- Collection (elements are scalar/enum) -----

    @PxlColumn(name = "StringList", exportSample = "Apple;Banana;Cherry")
    private List<String> stringList;

    @PxlColumn(name = "IntList", exportSample = "10;20;30")
    private List<Integer> intList;

    @PxlColumn(name = "GradeList", exportSample = "A;B;F")
    private List<Grade> gradeList;

}
