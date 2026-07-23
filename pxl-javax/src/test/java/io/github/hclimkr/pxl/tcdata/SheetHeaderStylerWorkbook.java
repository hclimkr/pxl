package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import io.github.hclimkr.pxl.styler.header.PxlHeaderHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderWrapTextStyler;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying sheet-level header stylers (exportSheetRequired/OptionalHeaderCellStyler).
 * Required header = horizontal center, optional header = wrap text.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SheetHeaderStylerWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Header",
            exportSheetRequiredHeaderCellStyler = PxlHeaderHorizontalCenterTextStyler.class,
            exportSheetOptionalHeaderCellStyler = PxlHeaderWrapTextStyler.class)
    private List<HeaderStyleRow> rows;

}
