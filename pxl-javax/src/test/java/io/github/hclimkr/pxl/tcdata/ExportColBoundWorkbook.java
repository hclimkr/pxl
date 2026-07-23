package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * exportLastDataColumnIndex is smaller than the actual column count, so the export must throw. (Employee has 7 columns)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExportColBoundWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "S", exportLastDataColumnIndex = 1)
    private List<Employee> rows;

}
