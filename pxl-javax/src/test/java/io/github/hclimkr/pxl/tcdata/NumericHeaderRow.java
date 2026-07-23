package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Verifies that a numeric header cell (2024) matches a string column name ("2024").
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NumericHeaderRow {

    @PxlColumn(name = "2024")
    private String year;

}
