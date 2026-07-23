package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import io.github.hclimkr.pxl.styler.data.PxlDataHorizontalCenterTextStyler;
import lombok.*;

import java.util.List;

/**
 * Verifies sheet-level data cell styler (exportSheetDataCellStyler).
 * When a column has no styler of its own, the sheet styler cascades down.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SheetStylerWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Centered", exportSheetDataCellStyler = PxlDataHorizontalCenterTextStyler.class)
    private List<Employee> rows;

}
