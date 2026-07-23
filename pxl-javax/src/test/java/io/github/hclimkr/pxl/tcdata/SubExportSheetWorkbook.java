package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying sheet override on export.
 * With exportOverrideSuperClassSheet=true, it overrides the super's same-named "Employees" sheet, so only the sub's data is exported.
 */
public class SubExportSheetWorkbook extends SuperExportSheetWorkbook {

    @PxlSheet(name = "Employees", exportOverrideSuperClassSheet = true)
    public List<Employee> employees;

}
