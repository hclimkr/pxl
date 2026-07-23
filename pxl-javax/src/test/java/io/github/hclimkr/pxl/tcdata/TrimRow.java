package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO with a column that trims strings on export (exportTrim).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TrimRow {

    @PxlColumn(name = "Padded", exportTrim = true)
    private String padded;

}
