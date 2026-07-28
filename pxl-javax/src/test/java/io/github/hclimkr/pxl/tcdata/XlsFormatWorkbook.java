package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.PxlFileFormat;
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import lombok.*;

import java.util.List;

/**
 * Workbook with a class-level @PxlWorkbook(exportFileFormat=HSSF).
 * Used to verify that PxlFileFormat.fromWorkbookObject reads the annotation value.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportFileFormat = PxlFileFormat.HSSF)
public class XlsFormatWorkbook {

    @PxlSheet(name = "Employees")
    private List<Employee> employees;

}
