package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import io.github.hclimkr.pxl.styler.data.PxlDataHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderWrapTextStyler;
import lombok.*;

import java.util.List;

/**
 * Workbook that names all three cell stylers on the annotation rather than in a runtime option.
 * Neither the sheet nor the columns name one, so the workbook level is what the cells end up with.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(
        exportWorkbookRequiredHeaderCellStyler = PxlHeaderHorizontalCenterTextStyler.class,
        exportWorkbookOptionalHeaderCellStyler = PxlHeaderWrapTextStyler.class,
        exportWorkbookDataCellStyler = PxlDataHorizontalCenterTextStyler.class
)
public class WorkbookStylerWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Styled")
    private List<HeaderStyleRow> rows;

}
