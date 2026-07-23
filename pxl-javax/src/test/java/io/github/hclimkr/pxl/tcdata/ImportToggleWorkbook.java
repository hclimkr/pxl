package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying sheet-level importEnabled.
 * enabled is imported, while disabled (importEnabled=false) is not imported even if the sheet exists.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ImportToggleWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Enabled")
    private List<Employee> enabled;

    @PxlSheet(name = "Disabled", importEnabled = false)
    private List<Employee> disabled;

}
