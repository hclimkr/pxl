package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Superclass workbook that declares the override itself, for verifying the direction of
 * importOverrideSuperClassSheet. (public fields for access, no Lombok)
 */
public class SuperOverrideCompanyWorkbook {

    @PxlSheet(name = "Departments", importOverrideSuperClassSheet = true)
    public List<Department> departments;

}
