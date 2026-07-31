package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;

import java.util.List;

/**
 * Superclass workbook that declares the override itself, for verifying the direction of
 * exportOverrideSuperClassSheet. (public fields for access, no Lombok)
 */
public class SuperOverrideExportSheetWorkbook {

    @PxlWorkbookName
    public String workbookName;

    @PxlSheet(name = "Employees", exportOverrideSuperClassSheet = true)
    public List<Employee> employees;

}
