package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that a sheet excluded from import claims no override.
 * <p>
 * departments names the super's "Departments" sheet and asks to override it, but importEnabled=false, so the name is
 * never claimed: the super field is bound as usual and this field stays null.
 */
public class SubDisabledOverrideCompanyWorkbook extends SuperCompanyWorkbook {

    @PxlSheet(name = "Departments", importEnabled = false, importOverrideSuperClassSheet = true)
    public List<Department> departments;

}
