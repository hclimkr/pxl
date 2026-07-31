package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that shadowing a field is not by itself a sheet override on export.
 * <p>
 * employees shadows the super field of the same name but declares a different sheet name ("Staff"), so nothing is
 * overridden and the workbook carries both sheets.
 */
public class SubOtherNameOverrideExportSheetWorkbook extends SuperExportSheetWorkbook {

    @PxlSheet(name = "Staff", exportOverrideSuperClassSheet = true)
    public List<Employee> employees;

}
