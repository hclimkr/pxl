package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.time.*;
import java.util.Date;

/**
 * DTO holding every date/time type with NO {@code pattern} specified on any column.
 * <p>
 * Without a pattern (nor masking), each date/time column is exported as a numeric Excel-date cell rather than a string,
 * so this fixture drives the "written as numeric date cells" export verification. (AllTypesRow now pins an explicit
 * pattern on its date/time columns, which forces string cells, so it can no longer serve that scenario.)
 * <p>
 * All values use only second granularity so they round-trip through the numeric serial (which preserves neither
 * nanoseconds nor offset/zone) unchanged.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class DateTimeNumericRow {

    @PxlColumn(name = "JavaDate", exportSample = "2023-06-15T10:30:45")
    private Date javaDate;

    @PxlColumn(name = "LocalDate", exportSample = "2023-06-15")
    private LocalDate localDate;

    @PxlColumn(name = "LocalTime", exportSample = "10:30:45")
    private LocalTime localTime;

    @PxlColumn(name = "LocalDateTime", exportSample = "2023-06-15T10:30:45")
    private LocalDateTime localDateTime;

    @PxlColumn(name = "ZonedDateTime", exportSample = "2023-06-15T10:30:45+09:00")
    private ZonedDateTime zonedDateTime;

    @PxlColumn(name = "OffsetTime", exportSample = "10:30:45+09:00")
    private OffsetTime offsetTime;

    @PxlColumn(name = "OffsetDateTime", exportSample = "2023-06-15T10:30:45+09:00")
    private OffsetDateTime offsetDateTime;

}
