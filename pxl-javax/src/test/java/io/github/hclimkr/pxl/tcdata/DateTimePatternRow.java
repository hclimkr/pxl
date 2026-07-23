package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.time.*;

/**
 * DTO with date/time fields each carrying a DateTimeFormatter pattern.
 * A pattern makes each value exported as a formatted text cell and re-parsed with the same formatter on import,
 * exercising each date/time codec's export-formatter and import-formatter branches on both directions.
 * (Values use second granularity so the patterns round-trip exactly.)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class DateTimePatternRow {

    @PxlColumn(name = "LocalDate", pattern = "yyyy/MM/dd")
    private LocalDate localDate;

    @PxlColumn(name = "LocalTime", pattern = "HH:mm:ss")
    private LocalTime localTime;

    @PxlColumn(name = "LocalDateTime", pattern = "yyyy/MM/dd HH:mm:ss")
    private LocalDateTime localDateTime;

    @PxlColumn(name = "OffsetTime", pattern = "HH:mm:ssXXX")
    private OffsetTime offsetTime;

    @PxlColumn(name = "OffsetDateTime", pattern = "yyyy-MM-dd'T'HH:mm:ssXXX")
    private OffsetDateTime offsetDateTime;

}
