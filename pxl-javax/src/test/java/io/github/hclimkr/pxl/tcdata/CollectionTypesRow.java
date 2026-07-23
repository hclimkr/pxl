package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Date;
import java.util.List;

/**
 * DTO whose columns are collections of many different element types.
 * <p>
 * Round-tripping this exercises every element-type branch of the Collection codec (parse and build),
 * and along the way each element's wrapper/date-time codec string-parse and string-build path.
 * (String/Integer/enum element collections are already covered by AllTypesRow, so they are omitted here.)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CollectionTypesRow {

    @PxlColumn(name = "Bytes")
    private List<Byte> bytes;

    @PxlColumn(name = "Shorts")
    private List<Short> shorts;

    @PxlColumn(name = "Longs")
    private List<Long> longs;

    @PxlColumn(name = "Doubles")
    private List<Double> doubles;

    @PxlColumn(name = "Floats")
    private List<Float> floats;

    @PxlColumn(name = "Chars")
    private List<Character> chars;

    @PxlColumn(name = "Bools")
    private List<Boolean> bools;

    @PxlColumn(name = "BigInts")
    private List<BigInteger> bigInts;

    @PxlColumn(name = "BigDecs")
    private List<BigDecimal> bigDecs;

    @PxlColumn(name = "LocalDates")
    private List<LocalDate> localDates;

    @PxlColumn(name = "LocalTimes")
    private List<LocalTime> localTimes;

    @PxlColumn(name = "LocalDateTimes")
    private List<LocalDateTime> localDateTimes;

    @PxlColumn(name = "ZonedDateTimes")
    private List<ZonedDateTime> zonedDateTimes;

    @PxlColumn(name = "OffsetTimes")
    private List<OffsetTime> offsetTimes;

    @PxlColumn(name = "OffsetDateTimes")
    private List<OffsetDateTime> offsetDateTimes;

    @PxlColumn(name = "JavaDates")
    private List<Date> javaDates;

    @PxlColumn(name = "Durations")
    private List<Duration> durations;

    @PxlColumn(name = "Periods")
    private List<Period> periods;

    @PxlColumn(name = "Moneys")
    private List<Money> moneys;

}
