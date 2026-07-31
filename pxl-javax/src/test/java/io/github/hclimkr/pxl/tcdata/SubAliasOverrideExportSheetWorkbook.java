package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that a sheet override on export is decided by sheet name, not by field name.
 * <p>
 * The field is named staff rather than employees, and only the second candidate name overlaps the super's
 * "Employees" sheet - one overlapping candidate is enough, so the super field writes no sheet. The sheet this
 * field writes is named after the first candidate, "Crew".
 */
public class SubAliasOverrideExportSheetWorkbook extends SuperExportSheetWorkbook {

    @PxlSheet(name = {"Crew", "Employees"}, exportOverrideSuperClassSheet = true)
    public List<Employee> staff;

}
