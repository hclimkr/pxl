package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that a sheet excluded from export claims no override.
 * <p>
 * employees names the super's "Employees" sheet and asks to override it, but exportEnabled=false, so the name is
 * never claimed: the super field writes the sheet and this field writes nothing.
 */
public class SubDisabledOverrideExportSheetWorkbook extends SuperExportSheetWorkbook {

    @PxlSheet(name = "Employees", exportEnabled = false, exportOverrideSuperClassSheet = true)
    public List<Employee> employees;

}
