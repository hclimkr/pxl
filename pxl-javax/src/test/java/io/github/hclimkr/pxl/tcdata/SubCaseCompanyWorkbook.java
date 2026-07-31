package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying that a superclass sheet override is recognized ignoring case.
 * <p>
 * departments names the super's "Departments" sheet in a different case ("DEPARTMENTS") with
 * importOverrideSuperClassSheet=true, so only the sub field is bound and the super field stays null.
 */
public class SubCaseCompanyWorkbook extends SuperCompanyWorkbook {

    @PxlSheet(name = "DEPARTMENTS", importOverrideSuperClassSheet = true)
    public List<Department> departments;

}
