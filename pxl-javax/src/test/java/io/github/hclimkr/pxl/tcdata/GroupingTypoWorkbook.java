package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook specifying a non-existent field name in exportGroupingFieldName. Must fail-fast on export.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class GroupingTypoWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Employees", exportGroupingFieldName = "noSuchField")
    private List<Employee> employees;

}
