package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * A @PxlWorkbook-style DTO. Has a workbook name field and two sheets (with different row types).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CompanyWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Employees")
    private List<Employee> employees;

    @PxlSheet(name = "Departments")
    private List<Department> departments;

}
