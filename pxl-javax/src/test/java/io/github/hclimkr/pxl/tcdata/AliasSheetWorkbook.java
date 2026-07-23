package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying that the actual sheet name matches any of the multiple sheet names (name={...}) specified.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class AliasSheetWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = {"Crew", "Employee", "직원"})
    private List<Employee> data;

}
