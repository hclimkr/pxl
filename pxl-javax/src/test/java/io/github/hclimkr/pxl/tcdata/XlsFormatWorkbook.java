package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import lombok.*;

import java.util.List;

/**
 * Workbook with a class-level @PxlWorkbook(exportExcelEngine=HSSF), which writes the XLS format.
 * Used to verify that PxlExcelEngine.fromWorkbookObject reads the annotation value.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportExcelEngine = PxlExcelEngine.HSSF)
public class XlsFormatWorkbook {

    @PxlSheet(name = "Employees")
    private List<Employee> employees;

}
