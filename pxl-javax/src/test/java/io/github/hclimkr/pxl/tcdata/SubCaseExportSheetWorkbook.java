package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that a sheet override on export is recognized ignoring case.
 * <p>
 * employees names the super's "Employees" sheet in a different case ("EMPLOYEES") with
 * exportOverrideSuperClassSheet=true, so only the sub's data is exported, as a single sheet.
 */
public class SubCaseExportSheetWorkbook extends SuperExportSheetWorkbook {

    @PxlSheet(name = "EMPLOYEES", exportOverrideSuperClassSheet = true)
    public List<Employee> employees;

}
