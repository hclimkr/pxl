package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying sheet header/data row and column index (export/import) combinations.
 * export writes at the specified positions, and import must read from the same positions to round-trip. (1-based)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class IndexShiftWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Data",
            exportHeaderRowIndex = 3, exportFirstDataRowIndex = 4, exportFirstDataColumnIndex = 2,
            importHeaderRowIndex = 3, importFirstDataRowIndex = 4, importFirstDataColumnIndex = 2)
    private List<Employee> data;

}
