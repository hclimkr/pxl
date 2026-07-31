package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass of a workbook whose superclass carries importOverrideSuperClassSheet.
 * <p>
 * The flag points from a subclass toward its superclass only, so the super field cannot suppress this one:
 * both fields bind the same "Departments" sheet.
 */
public class SubOverrideCompanyWorkbook extends SuperOverrideCompanyWorkbook {

    @PxlSheet(name = "Departments")
    public List<Department> departments;

}
