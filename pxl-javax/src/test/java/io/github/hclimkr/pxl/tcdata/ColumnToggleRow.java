package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying per-column export/import enable toggles.
 * <ul>
 *   <li>always: both enabled (default)</li>
 *   <li>exportOff: exportEnabled=false -> the column itself is not created on export</li>
 *   <li>importOff: importEnabled=false -> not imported even if the header exists (always null)</li>
 * </ul>
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ColumnToggleRow {

    @PxlColumn(name = "Always")
    private String always;

    @PxlColumn(name = "ExportOff", exportEnabled = false)
    private String exportOff;

    @PxlColumn(name = "ImportOff", importEnabled = false)
    private String importOff;

}
