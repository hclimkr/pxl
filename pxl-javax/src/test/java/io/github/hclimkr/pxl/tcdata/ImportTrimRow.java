package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying importTrim combinations. raw is not trimmed (whitespace preserved), trimmed uses the default (trim).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ImportTrimRow {

    @PxlColumn(name = "Raw", importTrim = false)
    private String raw;

    @PxlColumn(name = "Trimmed")   // importTrim defaults to true
    private String trimmed;

}
