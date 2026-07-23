package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.time.LocalDateTime;

/**
 * DTO for verifying the 1904 date-windowing branch split in the LocalDateTime codec.
 * <p>
 * Both columns hold the identical raw NUMERIC serial, but "Formatted" is date-formatted while "Plain" is a bare
 * number: the codec reads the former via POI's workbook-aware {@code getLocalDateTimeCellValue()} (honors the
 * workbook's 1904 date system) and the latter via {@code DateUtil.getLocalDateTime} (fixed to the 1900 system).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class DateWindowingRow {

    // Date-formatted numeric cell -> isCellDateFormatted == true -> getLocalDateTimeCellValue() (honors workbook 1904 windowing)
    @PxlColumn(name = "Formatted")
    private LocalDateTime formatted;

    // Plain numeric cell (same raw serial, no date format) -> isCellDateFormatted == false -> DateUtil.getLocalDateTime (1900 windowing)
    @PxlColumn(name = "Plain")
    private LocalDateTime plain;

}
