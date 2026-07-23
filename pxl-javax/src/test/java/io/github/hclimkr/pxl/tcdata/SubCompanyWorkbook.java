package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass workbook for verifying superclass sheet inheritance/override.
 * <p>
 * employees is not overridden, so both the super and sub fields are bound;
 * departments has importOverrideSuperClassSheet=true, so only the sub field is bound and the super field stays null.
 */
public class SubCompanyWorkbook extends SuperCompanyWorkbook {

    @PxlSheet(name = "Employees")
    public List<Employee> employees;

    @PxlSheet(name = "Departments", importOverrideSuperClassSheet = true)
    public List<Department> departments;

}
