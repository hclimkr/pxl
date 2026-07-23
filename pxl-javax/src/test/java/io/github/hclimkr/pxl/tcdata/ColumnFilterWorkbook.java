package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying exportColumnFilter (auto filter).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ColumnFilterWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Filtered", exportColumnFilter = true)
    private List<Employee> rows;

}
