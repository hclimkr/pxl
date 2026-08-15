package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook that asks to skip hidden rows and columns through the annotation rather than a runtime sheet option.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class HiddenSheetWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Data", importExcludeHiddenRows = true, importExcludeHiddenColumns = true)
    private List<Employee> rows;

}
