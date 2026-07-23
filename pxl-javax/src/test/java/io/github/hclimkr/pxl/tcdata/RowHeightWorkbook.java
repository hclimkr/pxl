package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying exportRowHeightInPoints (row height).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class RowHeightWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Tall", exportRowHeightInPoints = 40)
    private List<Employee> rows;

}
