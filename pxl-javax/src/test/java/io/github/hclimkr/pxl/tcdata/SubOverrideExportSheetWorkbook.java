package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import java.util.List;

/**
 * Subclass of a workbook whose superclass carries exportOverrideSuperClassSheet.
 * <p>
 * The flag points from a subclass toward its superclass only, so the super field cannot suppress this one: both
 * fields ask for an "Employees" sheet, which a workbook cannot hold twice.
 */
public class SubOverrideExportSheetWorkbook extends SuperOverrideExportSheetWorkbook {

    @PxlSheet(name = "Employees")
    public List<Employee> employees;

}
