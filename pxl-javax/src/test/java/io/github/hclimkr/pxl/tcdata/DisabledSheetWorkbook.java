package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * A workbook whose only sheet has exportEnabled=false.
 * Since there is no sheet to export (0 sheets), the export must fail.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class DisabledSheetWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Disabled", exportEnabled = false, exportSampleEnabled = false)
    private List<Employee> rows;

}
