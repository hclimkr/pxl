package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import lombok.*;

import java.util.List;

/**
 * Workbook that picks the streaming writer and sizes its row-access window on the annotation.
 * <p>
 * The window is deliberately smaller than the row count the test writes, so rows leave memory while the sheet is
 * still being built - the case a window wider than the data never reaches.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportExcelEngine = PxlExcelEngine.SXSSF, exportSXSSFRowAccessWindowSize = 10)
public class SxssfWindowWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "People")
    private List<Employee> people;

}
