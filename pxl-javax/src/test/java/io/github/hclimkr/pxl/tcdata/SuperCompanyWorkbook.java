package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Superclass workbook for verifying superclass sheet inheritance/override.
 * <p>
 * Fields are public for access, and Lombok getters are not used to avoid name collisions with the subclass.
 */
public class SuperCompanyWorkbook {

    @PxlSheet(name = "Employees")
    public List<Employee> employees;

    @PxlSheet(name = "Departments")
    public List<Department> departments;

}
