package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.time.Duration;
import java.time.Period;

/**
 * DTO for verifying Duration/Period import with a custom DurationFormatUtils-style importPattern.
 * The pattern is a column-level constant compiled once per column and reused across rows; a value that
 * does not match the pattern falls back to ISO-8601 parsing.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TemporalPatternRow {

    @PxlColumn(name = "Dur", importPattern = "HH:mm:ss")
    private Duration dur;

    @PxlColumn(name = "Per", importPattern = "yy/MM/dd")
    private Period per;

}
