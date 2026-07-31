package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that an override is decided by sheet name, not by field name.
 * <p>
 * The field is named depts rather than departments, and only the second candidate name overlaps the super's
 * "Departments" sheet - one overlapping candidate is enough, so the super field is suppressed and stays null.
 */
public class SubAliasOverrideCompanyWorkbook extends SuperCompanyWorkbook {

    @PxlSheet(name = {"Divisions", "Departments"}, importOverrideSuperClassSheet = true)
    public List<Department> depts;

}
