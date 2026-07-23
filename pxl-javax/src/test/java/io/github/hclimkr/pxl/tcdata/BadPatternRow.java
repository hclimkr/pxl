package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.time.LocalDate;

/**
 * (Guard regression) Verifies that an invalid date pattern leads to an exception at build time.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadPatternRow {

    @PxlColumn(name = "Date", pattern = "yyyy'")   // unclosed quote literal -> invalid pattern
    private LocalDate date;

}
