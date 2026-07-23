package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Custom object column with an incorrectly declared converter.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadConverterRow {

    @PxlColumn(name = "Bad")
    private BadExportConverterObject bad;

}
