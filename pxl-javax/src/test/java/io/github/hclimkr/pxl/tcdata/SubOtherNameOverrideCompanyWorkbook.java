package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that shadowing a field is not by itself an override.
 * <p>
 * departments shadows the super field of the same name but declares a different sheet name ("Divisions"), so the
 * super's "Departments" sheet is bound as usual and this field stays null - no such sheet exists in the source.
 */
public class SubOtherNameOverrideCompanyWorkbook extends SuperCompanyWorkbook {

    @PxlSheet(name = "Divisions", importOverrideSuperClassSheet = true)
    public List<Department> departments;

}
