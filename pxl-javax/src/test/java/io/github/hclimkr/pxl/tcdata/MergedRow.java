package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Row DTO for verifying merged-cell (importEachCellOfMergedRegion) import.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class MergedRow {

    @PxlColumn(name = "Region")
    private String region;

    @PxlColumn(name = "Terminal")
    private String terminal;

    @PxlColumn(name = "Stop")
    private String stop;

    @PxlColumn(name = "Destination")
    private String destination;

}
