package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;

import java.util.List;

/**
 * Superclass workbook for verifying sheet override on export. (public fields for access, no Lombok)
 */
public class SuperExportSheetWorkbook {

    @PxlWorkbookName
    public String workbookName;

    @PxlSheet(name = "Employees")
    public List<Employee> employees;

}
