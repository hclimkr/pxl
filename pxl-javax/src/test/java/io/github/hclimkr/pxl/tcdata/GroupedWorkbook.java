package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying grouping export. Split into multiple sheets by Employee.department value.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class GroupedWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Employees", exportGroupingFieldName = "department")
    private List<Employee> employees;

}
