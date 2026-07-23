package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

/**
 * DTO for verifying custom-pattern (pattern / importPattern / exportPattern) combinations.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class PatternRow {

    // pattern applies to both import and export.
    @PxlColumn(name = "Date", pattern = "yyyy/MM/dd")
    private LocalDate date;

    @PxlColumn(name = "Time", pattern = "HH.mm.ss")
    private LocalTime time;

    // Number format (DecimalFormat) pattern - thousands separator + 2 decimal places
    @PxlColumn(name = "Amount", pattern = "#,##0.00")
    private BigDecimal amount;

    // Specify importPattern/exportPattern separately (same pattern)
    @PxlColumn(name = "Timestamp", importPattern = "yyyy.MM.dd HH:mm", exportPattern = "yyyy.MM.dd HH:mm")
    private LocalDateTime timestamp;

}
